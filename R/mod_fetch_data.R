#' fetch_data UI Function
#'
#' @description A shiny Module.
#'
#' @param id,input,output,session Internal parameters for {shiny}.
#'
#' @noRd
#'
mod_fetch_data_ui <- function(id) {
  ns <- shiny::NS(id)
  shiny::tagList()
}

#' fetch_data Server Functions
#'
#' @noRd
mod_fetch_data_server <- function(id) {
  shiny::moduleServer(id, function(input, output, session) {
    ns <- session$ns
  })
}

## To be copied in the UI
# mod_fetch_data_ui("fetch_data_1")

## To be copied in the server
# mod_fetch_data_server("fetch_data_1")
