# Clean Excel files -----
## Functions -----
### Extract Excel cells ----
extract_cells <- function(file_path, sheet_name) {
  # Convert each Excel cell into a row with metadata
  cells <- tidyxl::xlsx_cells(
    file_path,
    sheet = sheet_name
  ) |>
    dplyr::filter(!is_blank) |>
    dplyr::select(
      "address",
      "row",
      "col",
      "data_type",
      "character",
      "numeric",
      "local_format_id"
    )
  cells
}

# Clean variables
clean_names <- function(variable) {
  variable <- plyr::mapvalues(
    variable,
    from = unique(variable),
    to = janitor::make_clean_names(unique(variable))
  )
  variable
}

# Get products per hierarchy level
get_hierarchy_level <- function(data, level) {
  variable_name <- paste0("category_", level)
  data |>
    dplyr::filter(.data$hierarchy == level) |>
    dplyr::select("row", "col", {{ variable_name }} := "product")
}

# Use the first sheet as template to extract the hierarchies
get_hierarchy_table <- function(file_path) {
  # Get personalized Excel formats, that convey meaning of hierarchy
  formats <- tidyxl::xlsx_formats(file_path)
  indent <- formats$local$alignment$indent # Maps each local_format_id to an indentation number
  sheet_names <- readxl::excel_sheets(file_path)
  # If there are 13 sheets, the first is a front page and should be skipped
  if (length(sheet_names) == 13) {
    sheet_index <- 2
  } else {
    sheet_index <- 1
  }
  cells <- extract_cells(file_path, sheet_names[sheet_index]) |>
    dplyr::select("row", "col", product = "character", "local_format_id")

  # Get each product and its hierarchy
  products_hierarchy <- cells |>
    dplyr::filter(.data$col == 1, .data$row >= 4) |>
    dplyr::mutate(
      hierarchy = indent[.data$local_format_id], # 0 is at the top, then 1, etc.
      product = clean_names(.data$product)
    )
  hierarchy_0 <- get_hierarchy_level(products_hierarchy, 0)
  hierarchy_1 <- get_hierarchy_level(products_hierarchy, 1)
  hierarchy_2 <- get_hierarchy_level(products_hierarchy, 2)
  hierarchy_3 <- get_hierarchy_level(products_hierarchy, 3)
  hierarchy_4 <- get_hierarchy_level(products_hierarchy, 4)
  hierarchy_5 <- get_hierarchy_level(products_hierarchy, 5)

  # Add nested hierarchies
  products_0 <- products_hierarchy |>
    dplyr::filter(.data$hierarchy == 0)
  products_1 <- products_hierarchy |>
    dplyr::filter(.data$hierarchy == 1) |>
    unpivotr::enhead(hierarchy_0, "left-up")
  products_2 <- products_hierarchy |>
    dplyr::filter(.data$hierarchy == 2) |>
    unpivotr::enhead(hierarchy_0, "left-up") |>
    unpivotr::enhead(hierarchy_1, "left-up")
  products_3 <- products_hierarchy |>
    dplyr::filter(.data$hierarchy == 3) |>
    unpivotr::enhead(hierarchy_0, "left-up") |>
    unpivotr::enhead(hierarchy_1, "left-up") |>
    unpivotr::enhead(hierarchy_2, "left-up")
  products_4 <- products_hierarchy |>
    dplyr::filter(.data$hierarchy == 4) |>
    unpivotr::enhead(hierarchy_0, "left-up") |>
    unpivotr::enhead(hierarchy_1, "left-up") |>
    unpivotr::enhead(hierarchy_2, "left-up") |>
    unpivotr::enhead(hierarchy_3, "left-up")
  products_5 <- products_hierarchy |>
    dplyr::filter(.data$hierarchy == 5) |>
    unpivotr::enhead(hierarchy_0, "left-up") |>
    unpivotr::enhead(hierarchy_1, "left-up") |>
    unpivotr::enhead(hierarchy_2, "left-up") |>
    unpivotr::enhead(hierarchy_3, "left-up") |>
    unpivotr::enhead(hierarchy_4, "left-up")

  # Bind all products
  data_hierarchy <- dplyr::bind_rows(
    products_0,
    products_1,
    products_2,
    products_3,
    products_4,
    products_5
  ) |>
    dplyr::select(-c("col", "local_format_id")) |>
    # There are duplicated names in different subcategories, so we keep row as identifier
    # Example, three rows for PIMIENTOS (in T.HORTALIZAS FRESCAS, FRUTA&HORTA.CONSERVA, FRUTA&HORTA.CONGELAD)
    dplyr::mutate(
      parent_category = dplyr::coalesce(
        .data$category_4,
        .data$category_3,
        .data$category_2,
        .data$category_1,
        .data$category_0
      )
    )
  data_hierarchy
}

# Function to check for a leap year
is_leap_year <- function(year) {
  if ((year %% 4 == 0 && year %% 100 != 0) || year %% 400 == 0) {
    return(TRUE)
  } else {
    return(FALSE)
  }
}

### Transform Excel cells ----
clean_file <- function(file_path) {
  # Loop for all sheets
  sheet_names <- readxl::excel_sheets(file_path)
  # Some files have a first page without data
  if (length(sheet_names) == 13) {
    sheet_names <- sheet_names[-1]
  }
  data_list <- vector(mode = "list", length = length(sheet_names))
  file_name <- stringr::str_extract(file_path, "[0-9]{4}.*$")
  year <- as.numeric(substr(file_name, 1, 4))
  census_base <- stringr::str_extract(file_name, "base([0-9]{4})", group = 1)
  message(
    "Cleaning file: ",
    file_name,
    ", number of sheets: ",
    length(sheet_names)
  )
  for (sheet_index in seq_along(sheet_names)) {
    cells <- extract_cells(file_path, sheet_names[sheet_index])

    # Subset headers for joining later
    cells_products <- cells |>
      dplyr::filter(.data$col == 1, .data$row >= 4) |>
      dplyr::select("row", "col", product = "character")

    # Get the values
    values <- cells |>
      dplyr::filter(.data$row >= 2, col >= 2) |>
      unpivotr::behead(direction = "NNW", name = "region") |> # finds the region as a header direction north-west (up, then left)
      unpivotr::behead(direction = "N", name = "variable") |> # the variable is right above the cell value (north: up)
      unpivotr::enhead(cells_products, direction = "left") |> # adds back the product name
      dplyr::select(
        "row",
        "variable",
        "region",
        "product",
        "value" = "numeric"
      ) |>
      dplyr::mutate(
        year = year,
        month = sheet_index,
        date = lubridate::make_date(.data$year, .data$month, 1L)
      )

    message(
      "Cleaned file: ",
      file_name,
      ", sheet: ",
      sheet_index,
      "/",
      length(sheet_names)
    )
    data_list[[sheet_index]] <- values
  }
  message("Joining months ...")
  data <- dplyr::bind_rows(data_list) |>
    # Clean names
    dplyr::mutate(
      census_base = census_base,
      dplyr::across(
        tidyselect::all_of(c("product", "variable", "region")),
        clean_names
      )
    ) |>
    # Pivot variables to make plotting easier
    tidyr::pivot_wider(
      values_from = "value",
      names_from = "variable",
      id_cols =
    ) |>
    # Adjust for different number of days per month
    dplyr::mutate(
      days_month = dplyr::case_when(
        .data$month %in% c(1, 3, 5, 7, 8, 10, 12) ~ 31,
        .data$month %in% c(4, 6, 9, 11) ~ 30,
        .data$month %in% c(2) & !is_leap_year(.env$year) ~ 28,
        .data$month %in% c(2) & is_leap_year(.env$year) ~ 29
      ),
      consumo_capita_dia = .data$consumo_x_capita / days_month,
      gasto_capita_dia = .data$gasto_x_capita / days_month
    ) |>
    # Get duplicated product name
    dplyr::group_by(.data$product, .data$month, .data$region) |>
    dplyr::mutate(duplicated = dplyr::n()) |>
    dplyr::ungroup()

  # Obtain nested hierarchy
  # This is only available from 2014 onwards, before that there's no
  # explicit hierarchy as cell indentation
  if (year >= 2014) {
    message("Adding nested hierarchy to ", file_name, " ...")
    hierarchy <- get_hierarchy_table(file_path)
    data <- data |>
      dplyr::left_join(hierarchy, by = c("product", "row"))
  }
  message("Finished cleaning file: ", file_name, " -----------------------")
  data
}

# Step by step
file_dir <- here::here("data-raw")
file_list <- list.files(file_dir, pattern = "*.xlsx")
loop_over <- seq_along(file_list)
# loop_over <- length(file_list)
data_list <- vector(mode = "list", length = length(loop_over))
for (i in loop_over) {
  file_name <- file_list[i]
  file_path <- here::here(file_dir, file_name)
  data_list[[i]] <- clean_file(file_path)
}
(data <- dplyr::bind_rows(data_list))

data |>
  dplyr::filter(
    region == "t_espana",
    product == "platos_preparados"
  ) |>
  ggplot2::ggplot(
    ggplot2::aes(
      x = date,
      y = consumo_x_capita,
      group = census_base,
      color = as.factor(month)
    )
  ) +
  ggplot2::geom_line() +
  ggplot2::geom_point()

# Monthly data should account for the different numbers of days each month
# (February always have less total consumption)

# Fit ARIMA models to account for seasonal change, and get yearly trends
# And seasonal profiles

# Adding unique ids for products/categories
data |>
  dplyr::filter(duplicated > 1) |>
  dplyr::mutate(
    full_name = dplyr::case_when(
      is.na(hierarchy) ~ paste("?", product, sep = "/"),
      hierarchy == 0 ~ product,
      hierarchy == 1 ~ paste(category_0, product, sep = "/"),
      hierarchy == 2 ~ paste(category_0, category_1, product, sep = "/"),
      hierarchy == 3 ~ paste(
        category_0,
        category_1,
        category_2,
        product,
        sep = "/"
      ),
      hierarchy == 4 ~ paste(
        category_0,
        category_1,
        category_2,
        category_3,
        product,
        sep = "/"
      ),
      hierarchy == 5 ~ paste(
        category_0,
        category_1,
        category_2,
        category_3,
        category_4,
        product,
        sep = "/"
      ),
    ),
    # Manual changes
    unique_name = dplyr::case_when(
      # frozen or conserved vegetables
      duplicated == 1 ~ product,
      stringr::str_detect(parent_category, "congelad") ~ paste0(
        product,
        "_congelados"
      ),
      stringr::str_detect(parent_category, "conserva") ~ paste0(
        product,
        "_conservas"
      ),
      stringr::str_detect(parent_category, "verd|fresca") ~ paste0(
        product,
        "_frescos"
      ),
      stringr::str_detect(parent_category, "fiambres") ~ paste0(
        product,
        "_fiambres"
      ),
      stringr::str_detect(parent_category, "x1_litro") ~ paste0(
        product,
        "_1_litro"
      ),
      stringr::str_detect(parent_category, "huevos") ~ paste0(
        "huevos_",
        product
      ),
      stringr::str_detect(parent_category, "cacao_soluble|arroz") ~ paste(
        parent_category,
        product,
        sep = "_"
      ),
      stringr::str_detect(parent_category, "mayor_1_litro") ~ paste(
        product,
        parent_category,
        sep = "_"
      ),
      .default = product
    )
  ) |>
  dplyr::count(
    product,
    # full_name,
    # duplicated,
    unique_name,
    parent_category,
    year,
    # census_base
  ) |>
  dplyr::filter(is.na(parent_category)) |>
  dplyr::arrange(
    product,
    dplyr::desc(year),
    # census_base,
    # duplicated,
    parent_category
  ) |>
  print(n = Inf)


# Something strange happens
# Alcachofas don't have precio_medio_kg
# December 2019 is different from the rest of the year
# This results in different rows per products, and some variables are not well read also
# I need to think a better way of identifying unique products

data |> dplyr::filter(year == 2019, product == "alcachofas")

products_2019_12 <- data |>
  dplyr::filter(year == 2019, month == 12) |>
  dplyr::distinct(product) |>
  dplyr::pull()
products_2019_11 <- data |>
  dplyr::filter(year == 2019, month == 11) |>
  dplyr::distinct(product) |>
  dplyr::pull()

products_2019_12[!products_2019_12 %in% products_2019_11]

# TODO: create hierarchy for each sheet with indentation (not only once for each book)
# If a sheet doesn't have indentation, apply the previous hierarchy

unique_product_names <- sort(unique(data$product))
length(unique_product_names)

# Some names are uninterpretable on their own (e.g., "x1_litro") and would need adding their parent category
# Products change names with time, making comparisons hard
data |>
  dplyr::count(product, year) |>
  dplyr::arrange(year) |>
  dplyr::mutate(n = "X") |>
  tidyr::pivot_wider(
    names_from = year,
    values_from = n,
    values_fill = "-"
  ) |>
  dplyr::arrange(product) |>
  print(n = Inf)
