#' @import rvg
#' @import miniUI
#' @import shiny
#' @export
rvg_gadget <- function() {

  ui <- miniPage(
    gadgetTitleBar("Send current plot to Microsoft documents"),
        miniContentPanel(
          fillCol(
            fillRow(
                    selectInput("format", "Format", choices = c("MS PowerPoint" = "pptx", "MS Word" = "docx"), selected = 1 ),
                    div(),
                    textInput( inputId = "basename", label = "File name", value = "Rplot" )
                    ),
            fillRow(
              downloadButton('downloadData', 'Download')
            ),
            fillRow(
              numericInput( "width", "Plot width", min = 3, value = 6, step = .5),
              numericInput( "height", "Plot height", min = 3, value = 6, step = .5),
              numericInput("pwidth", "Page width", min = 5, value = 10, step = .5),
              numericInput("pheight", "Page height", min = 5, value = 8, step = .5)
            ),
            fillRow(
              h4("Word only")
            ),
            fillRow(
              numericInput( "ml", "Margin left", min = 0, value = 1, step = .5),
              numericInput( "mr", "Margin right", min = 0, value = 1, step = .5),
              numericInput( "mt", "Margin top", min = 0, value = 1, step = .5),
              numericInput( "mb", "Margin bottom", min = 0, value = 1, step = .5)
            )
          )
        )
      )
  server <- function(input, output) {
    output$downloadData <- downloadHandler(
      filename = function() {
        paste0(input$basename, ".", input$format)
      },
      content = function(file) {
        plot_ <- try( recordPlot() )
        if( inherits(plot_, "try-error") ) stopApp()
        if (input$format == "docx"){
          write_docx(file = file, code = replayPlot(plot_),
                     width = input$width, height = input$height,
                     pagesize = c(width = input$pwidth, height = input$pheight),
                     margins = c(left = input$ml, right = input$mr,
                                 top = input$mt, bottom = input$mb)
                     )
        } else {
          write_pptx(file = file, code = replayPlot(plot_),
                     width = input$width, height = input$height,
                     size = c(width = input$pwidth, height = input$pheight))
        }
      }
    )

    observeEvent(input$done, {
      stopApp()
    })

  }

  plot_ <- try( recordPlot(), silent = TRUE )
  if( inherits(plot_, "try-error") ) {
    message("no plot found, exiting...")
    return(invisible())
  }
  runGadget(ui, server)
}

#' @title Export current plot to Word or PowerPoint document.
#'
#' @description Create a Microsoft Word or PowerPoint with the current plot
#' exported in a editable vector graphics format.
#'
#' @export
#' @examples
#' if (interactive()){
#'   plot(rnorm(100))
#'   rvg_addin()
#' }
rvg_addin <- function() {
  con <- rstudioapi::getActiveDocumentContext()
  text <- con$selection[[1]]$text

  rvg_gadget()
}

