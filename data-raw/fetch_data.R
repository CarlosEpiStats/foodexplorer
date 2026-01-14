# Define functions ----
# Shorthand for grep filtering of text
filter_grep <- function(text, filter, ...) {
  grep(filter, x = text, value = TRUE, ...)
}

# Fetch data files urls
fetch_urls <- function() {
  # Reach url
  url_base <- "https://www.mapa.gob.es"
  url_rel <- "/es/alimentacion/temas/consumo-tendencias/panel-de-consumo-alimentario/series-anuales/default.aspx"
  # Get html tree
  html <- rvest::read_html(paste0(url_base, url_rel))
  # Get links on the page body
  nodes <- rvest::html_nodes(html, xpath = "//body//a")
  links <- rvest::html_attr(nodes, "href")
  # Filter relevant links, as the file name is not consistent year after year
  relevant_links <- links |>
    filter_grep(".xlsx") |> # only xlsx files
    filter_grep("mensuales") |> # only monthly series
    filter_grep("-sd-|sociodemo", invert = TRUE) |> # remove -sd- files (sociodemographic)
    filter_grep("canales", invert = TRUE) # remove channel sales files
  # Obtain full url for each file
  urls <- paste0(url_base, relevant_links)
  urls
}

# Download the files
process_file <- function(url) {
  # Get the year of the file
  year <- gsub(".*/([0-9]{4}).*", "\\1", url) # year is a four-digit number preceded by a slash
  # Up to 2023, the 2011 census was used as base. From 2023 on, the 2021 census is used. I download both versions for comparison
  census_base <- ifelse(grepl("base2021", url, fixed = TRUE), 2021, 2011)
  message("Trying file, year: ", year, " census: ", census_base)
  file_name <- file.path(here::here(
    "data-raw",
    paste0(year, "_base", census_base, ".xlsx")
  ))
  file_downloaded <- download.file(url, file_name, mode = "wb") # mode="wb" needed to avoid file corruption
  if (file_downloaded == 0) {
    message("File downloaded at ", file_name)
    message("-------------------------")
  } else {
    stop("Problem downloading file at ", url)
  }
}

# Download the files -----
urls <- fetch_urls()
do.call(rbind, lapply(urls, process_file))
