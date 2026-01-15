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
  cells <- extract_cells(file_path, sheet_names[1]) |>
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
    dplyr::select(-c("col", "local_format_id"))
  # There are duplicated names in different subcategories, so we keep row as identifier
  # Example, three rows for PIMIENTOS (in T.HORTALIZAS FRESCAS, FRUTA&HORTA.CONSERVA, FRUTA&HORTA.CONGELAD)
  data_hierarchy
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
  year <- substr(file_name, 1, 4)
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
        "numeric"
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
      dplyr::across(
        tidyselect::all_of(c("product", "variable", "region")),
        clean_names
      )
    )

  # Obtain nested hierarchy
  message("Adding nested hierarchy to ", file_name, " ...")
  hierarchy <- get_hierarchy_table(file_path)
  data_hierarchy <- data |>
    dplyr::left_join(hierarchy, by = c("product", "row"))

  message("Finished cleaning file: ", file_name, " -----------------------")
  data_hierarchy
}

# Step by step
file_dir <- here::here("data-raw")
file_list <- list.files(file_dir, pattern = "*.xlsx")
# Test with two files
loop_over <- 28:29
data_list <- vector(mode = "list", length = length(loop_over))
for (i in loop_over) {
  file_name <- file_list[i]
  file_path <- here::here(file_dir, file_name)
  data_list[[i]] <- clean_file(file_path)
}
(data <- dplyr::bind_rows(data_list))
