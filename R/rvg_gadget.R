#' @import rvg
#' @import miniUI
#' @import shiny
rvg_gadget <- function() {

  ui <- miniPage(
    gadgetTitleBar("Send plot to MS document"),
    miniContentPanel(
      fillCol(
        fillRow(
          radioButtons("format", "Format",
                       choices = c("PowerPoint" = "pptx", "Word" = "docx", "Excel" = "xlsx"), inline = TRUE,
                       selected = "pptx" ),
          div(),
          textInput( inputId = "basename", label = "File name", value = "Rplot" ),
          flex = c(6,1,4) ),
        fillRow(
          div(style = "text-align:center;",
            downloadButton('downloadData', 'Download', class = "btn-primary") )
        ),
        fillRow(
          sliderInput( "width", "Plot width", min = 3, max = 20, value = 6, step = .5),
          div(),
          sliderInput( "height", "Plot height", min = 3, max = 20, value = 6, step = .5),
          flex = c(4,1,4) ),
        fillRow(
          sliderInput("pwidth", "Page width", min = 5, max = 20, value = 10, step = .5),
          div(),
          sliderInput("pheight", "Page height", min = 5, max = 20, value = 8, step = .5),
          flex = c(4,1,4) ),
        flex = c(2,2,2,2)
      )
    )
  )

  server <- function(input, output) {

    currplot <- reactiveValues(plot = recordPlot() )

    output$downloadData <- downloadHandler(
      filename = function() {
        paste0(input$basename, ".", input$format)
      },
      content = function(file) {
        if (input$format == "docx"){
          write_docx(file = file, code = replayPlot(currplot$plot),
                     width = input$width, height = input$height,
                     pagesize = c(width = input$pwidth, height = input$pheight)
                     )
        } else if (input$format == "pptx"){
          write_pptx(file = file, code = replayPlot(currplot$plot),
                     width = input$width, height = input$height,
                     size = c(width = input$pwidth, height = input$pheight))
        } else {
          write_xlsx(file = file, code = replayPlot(currplot$plot),
                     width = input$width, height = input$height,
                     size = c(width = input$pwidth, height = input$pheight))
        }
      }
    )

    observeEvent(input$done, {
      stopApp()
    })
    observeEvent(input$cancel, {
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

