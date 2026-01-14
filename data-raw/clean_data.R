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
### Transform Excel cells ----
transform_cells <- function() {}

# Step by step
file_dir <- here::here("data-raw")
files <- here::here(file_dir, list.files(file_dir, pattern = "*.xlsx"))
# Test with one file
file <- files[29]
sheet_names <- readxl::excel_sheets(file)
cells <- extract_cells(file, sheet_names[1])
# Get personalized Excel formats, that convey meaning of hierarchy
formats <- tidyxl::xlsx_formats(file)
indent <- formats$local$alignment$indent # Maps each local_format_id to an indentation number
# Get each product and its hierarchy
products_hierarchy <- cells |>
  dplyr::filter(.data$col == 1, .data$row >= 4) |>
  dplyr::select("row", "col", product = "character", "local_format_id") |>
  dplyr::mutate(hierarchy = indent[.data$local_format_id]) # 0 is at the top, then 1, etc.
# Store names of products and clean them
products_names_raw <- unique(products_hierarchy$product)
products_names_clean <- janitor::make_clean_names(products_names_raw)
# Apply the clean names back
products_hierarchy$product <- plyr::mapvalues(
  products_hierarchy$product,
  from = products_names_raw,
  to = products_names_clean
)
get_hierarchy_level <- function(data, level) {
  variable_name <- paste0("category_", level)
  data |>
    dplyr::filter(.data$hierarchy == level) |>
    dplyr::select("row", "col", {{ variable_name }} := "product")
}
hierarchy_0 <- get_hierarchy_level(products_hierarchy, 0)
hierarchy_1 <- get_hierarchy_level(products_hierarchy, 1)
