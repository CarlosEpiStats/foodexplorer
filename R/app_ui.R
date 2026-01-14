#' The application User-Interface
#'
#' @param request Internal parameter for `{shiny}`.
#'     DO NOT REMOVE.
#' @noRd
app_ui <- function(request) {
  shiny::tagList(
    # Leave this function for adding external resources
    golem_add_external_resources(),
    # Your application UI logic
    shiny::fluidPage(
      shinydashboard::dashboardPage(
        header = shinydashboard::dashboardHeader(
          title = "foodexplorer Shiny App"
        ),
        sidebar = shinydashboard::dashboardSidebar(
          shinydashboard::sidebarMenu(
            id = "tabs",
            shinydashboard::menuItem(
              "How to use",
              tabName = "tab-welcome",
              icon = shiny::icon("info")
            ),
            shinydashboard::menuItem(
              "Explore trends",
              tabName = "tab-trends",
              icon = shiny::icon("arrow-trend-up")
            ),
            shinydashboard::menuItem(
              "Download data",
              tabName = "tab-download",
              icon = shiny::icon("download")
            )
          )
        ),
        body = shinydashboard::dashboardBody()
      )
    )
  )
}

#' Add external Resources to the Application
#'
#' This function is internally used to add external
#' resources inside the Shiny application.
#'

#' @noRd
golem_add_external_resources <- function() {
  golem::add_resource_path(
    "www",
    app_sys("app/www")
  )

  shiny::tags$head(
    golem::favicon(),
    golem::bundle_resources(
      path = app_sys("app/www"),
      app_title = "foodexplorer"
    )
    # Add here other external resources
    # for example, you can add shinyalert::useShinyalert()
  )
}
