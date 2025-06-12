# ==============================================================================
# Description:
# A modern and robust Shiny app that uses a hybrid approach. The shiny.fluent
# UI is used for data display, while a standard Shiny action panel with a
# customized shinyFiles modal is used for a clean, reliable file selection.
# This version remembers the last-used folder and adds QC checklists to the
# exported Excel file.
#
# Required Packages:
# shiny, shiny.fluent, writexl, shinyFiles, dplyr, uuid, shinyjs
#
# To Run:
# 1. Make sure you have installed all the required packages:
#    install.packages(c("shiny", "shiny.fluent", "writexl", "shinyFiles", "dplyr", "uuid", "shinyjs"))
# 2. Save this script as 'app.R'.
# 3. Run the app by executing `shiny::runApp()` in the R console from the
#    directory where you saved the file.
# ==============================================================================

# Load necessary libraries
library(shiny)
library(shiny.fluent)
library(writexl)
library(shinyFiles)
library(dplyr)
library(uuid)
library(shinyjs)

# ==============================================================================
# UI Definition
# ==============================================================================

# Helper function to create a single "Effect" card UI
create_effect_card <- function(id) {
  include_opts <- list(list(key = "Y", text = "Y"), list(key = "N", text = "N"))
  effect_type_opts <- list(list(key = "Benefit", text = "Benefit"), list(key = "Risk", text = "Risk"))
  var_type_opts <- list(list(key = "Continuous", text = "Continuous"), list(key = "Binary", text = "Binary"))
  direction_opts <- list(list(key = "increase", text = "increase"), list(key = "decrease", text = "decrease"))
  
  div(
    class = "effect-card ms-depth-8",
    id = paste0("card_", id),
    Stack(
      tokens = list(childrenGap = 10),
      div(
        class = "card-header",
        Text(variant = "large", paste("Effect:", substr(id, 1, 8))),
        IconButton.shinyInput(paste0("delete_", id), iconProps = list(iconName = "Delete"), title = "Delete this effect")
      ),
      p("Click a field below, then use the 'Actions' panel to choose a file.", style="font-size:0.9em; color: #605e5c;"),
      Stack(
        horizontal = TRUE, tokens = list(childrenGap = 10),
        TextField.shinyInput(paste0("source_location_", id), label = "Source Location Path", value = "", readOnly=TRUE, class="path-input"),
        TextField.shinyInput(paste0("source_file_", id), label = "Source File Name", value = "", readOnly=TRUE, class="path-input")
      ),
      TextField.shinyInput(paste0("ards_path_", id), label = "ARDS File Path", value = "", readOnly=TRUE, class="path-input"),
      hr(),
      Stack(
        horizontal = TRUE, tokens = list(childrenGap = 10), verticalAlign = "end",
        Dropdown.shinyInput(paste0("include_", id), label = "Include", value = "Y", options = include_opts),
        TextField.shinyInput(paste0("effect_", id), label = "Effect Label", value = "New Endpoint"),
        Dropdown.shinyInput(paste0("effect_type_", id), label = "Effect Type", value = "Benefit", options = effect_type_opts)
      ),
      Stack(
        horizontal = TRUE, tokens = list(childrenGap = 10), verticalAlign = "end",
        Dropdown.shinyInput(paste0("variable_type_", id), label = "Variable Type", value = "Binary", options = var_type_opts),
        TextField.shinyInput(paste0("study_", id), label = "Study or Integration", value = "STUDY01"),
        TextField.shinyInput(paste0("population_", id), label = "Population", value = "All Patients")
      ),
      Stack(
        horizontal = TRUE, tokens = list(childrenGap = 10), verticalAlign = "end",
        TextField.shinyInput(paste0("timepoint_", id), label = "Timepoint", value = "12 Weeks"),
        Dropdown.shinyInput(paste0("direction_", id), label = "Improvement Direction", value = "increase", options = direction_opts)
      ),
      TextField.shinyInput(paste0("comments_", id), label = "Comments", multiline = TRUE, rows = 2),
      TextField.shinyInput(paste0("ards_filters_", id), label = "ARDS Filters", value = "ARM == 'TREATMENT'")
    )
  )
}

# Main UI definition
ui <- fluentPage(
  useShinyjs(), 
  tags$head(
    tags$script(HTML("
      $(document).on('click', '.path-input input', function() {
        var cardId = $(this).closest('.effect-card').attr('id');
        var inputType = 'other';
        if (this.id.includes('source_location') || this.id.includes('source_file')) {
          inputType = 'source';
        } else if (this.id.includes('ards_path')) {
          inputType = 'ards';
        }
        var info = { cardId: cardId, inputType: inputType };
        Shiny.setInputValue('last_clicked_card_info', info, { priority: 'event' });
      });
    ")),
    tags$style(HTML("
      body { background-color: #faf9f8; }
      .action-panel { padding: 15px; background-color: #f3f2f1; border: 1px solid #e1dfdd; border-radius: 4px; margin-top: 20px; }
      .action-panel .btn { width: 100%; margin-bottom: 10px; }
      .effect-card { margin-bottom: 20px; }
      #app-header { padding: 10px 20px; background-color: #ffffff; border-bottom: 1px solid #e1dfdd;}
      #main-content { padding: 20px; max-width: 1200px; margin: auto; }
      /* This CSS specifically hides the unwanted UI elements in the shinyFiles modal. */
      .sF-navigate, .sF-view, .sF-sort { 
        display: none !important; 
      }
    "))
  ),
  Stack(
    id = "app-header",
    horizontal = TRUE, tokens = list(childrenGap = 10), verticalAlign = "end",
    TextField.shinyInput("sheet_name", label = "Current Sheet Name", value = "Sheet1"),
    PrimaryButton.shinyInput("add_effect_btn", "Add New Effect", iconProps = list(iconName = "Add")),
    DefaultButton.shinyInput("add_sheet_btn", "Save as New Sheet")
  ),
  div(
    id = "main-content",
    div(id = "effect_cards_container"),
    
    div(class = "action-panel",
        h4("Actions"),
        p("Click a path field in a card above, then click the appropriate button below."),
        fluidRow(
          column(6, uiOutput("source_file_button_ui")),
          column(6, uiOutput("ards_file_button_ui"))
        ),
        hr(),
        fluidRow(
          column(12, downloadButton("global_download_btn", "Download All Sheets to Excel", class="btn btn-success btn-lg btn-block"))
        )
    ),
    
    hr(),
    Text(variant = "xxLarge", "Live Specification Preview"),
    uiOutput("live_data_table")
  )
)

# ==============================================================================
# Server Logic
# ==============================================================================
server <- function(input, output, session) {
  
  all_sheets <- reactiveVal(list())
  current_effects <- reactiveVal(list())
  active_card_info <- reactiveVal(NULL)
  volumes <- c(Home = fs::path_home(), "R Installation" = R.home(), getVolumes()())
  
  last_source_path <- reactiveVal(normalizePath("~"))
  last_ards_path <- reactiveVal(normalizePath("~"))
  
  # --- Dynamic UI for File Chooser Buttons ---
  output$source_file_button_ui <- renderUI({
    shinyFileChoose(input, "source_file_chooser", roots = volumes, session = session, defaultPath = last_source_path())
    shinyFilesButton("source_file_chooser", "Choose Source File", "Please select a source file", multiple=FALSE, class="btn btn-primary")
  })
  
  output$ards_file_button_ui <- renderUI({
    shinyFileChoose(input, "ards_file_chooser", roots = volumes, session = session, defaultPath = last_ards_path())
    shinyFilesButton("ards_file_chooser", "Choose ARDS File", "Please select an ARDS file", multiple=FALSE, class="btn btn-primary")
  })
  
  # --- Event Handlers ---
  observeEvent(input$add_effect_btn, {
    id <- UUIDgenerate()
    insertUI(selector = "#effect_cards_container", where = "beforeEnd", ui = create_effect_card(id))
    current_effects(c(isolate(current_effects()), id))
  })
  
  observeEvent(input$last_clicked_card_info, {
    active_card_info(input$last_clicked_card_info)
  })
  
  observeEvent(input$source_file_chooser, {
    info <- active_card_info()
    req(info, info$inputType == 'source', is.list(input$source_file_chooser))
    
    file_info <- parseFilePaths(volumes, input$source_file_chooser)
    if (nrow(file_info) > 0) {
      card_id_base <- sub("card_", "", info$cardId)
      updateTextField.shinyInput(session, paste0("source_location_", card_id_base), value = as.character(dirname(file_info$datapath)))
      updateTextField.shinyInput(session, paste0("source_file_", card_id_base), value = as.character(file_info$name))
      last_source_path(dirname(file_info$datapath))
    }
  })
  
  observeEvent(input$ards_file_chooser, {
    info <- active_card_info()
    req(info, info$inputType == 'ards', is.list(input$ards_file_chooser))
    
    file_info <- parseFilePaths(volumes, input$ards_file_chooser)
    if (nrow(file_info) > 0) {
      card_id_base <- sub("card_", "", info$cardId)
      updateTextField.shinyInput(session, paste0("ards_path_", card_id_base), value = as.character(file_info$datapath))
      last_ards_path(dirname(file_info$datapath))
    }
  })
  
  observe({
    req(current_effects())
    lapply(current_effects(), function(id) {
      observeEvent(input[[paste0("delete_", id)]], {
        removeUI(selector = paste0("#card_", id), immediate = TRUE)
        current_effects(isolate(current_effects())[isolate(current_effects()) != id])
      }, ignoreInit = TRUE, once = TRUE, autoDestroy = TRUE)
    })
  })
  
  # --- Reactive Data Frame for Live Preview ---
  live_df <- reactive({
    req(length(current_effects()) > 0)
    reactiveValuesToList(input)
    
    effect_ids <- isolate(current_effects())
    req(input[[paste0("include_", effect_ids[[1]])]])
    
    spec_list <- lapply(effect_ids, function(id) {
      data.frame(
        Include = input[[paste0("include_", id)]], Effect = input[[paste0("effect_", id)]],
        Effect_Type = input[[paste0("effect_type_", id)]], Variable_Type = input[[paste0("variable_type_", id)]],
        Study_or_Integration = input[[paste0("study_", id)]], Population = input[[paste0("population_", id)]],
        Timepoint = input[[paste0("timepoint_", id)]], Improvement_direction = input[[paste0("direction_", id)]],
        Comments = input[[paste0("comments_", id)]], Source_Location = input[[paste0("source_location_", id)]],
        Source_File = input[[paste0("source_file_", id)]], ARDS_Path = input[[paste0("ards_path_", id)]],
        ARDS_filters = input[[paste0("ards_filters_", id)]], stringsAsFactors = FALSE, check.names = FALSE
      )
    })
    bind_rows(spec_list)
  })
  
  output$live_data_table <- renderUI({
    df <- if (length(current_effects()) > 0) live_df() else data.frame()
    if (nrow(df) > 0) {
      DetailsList(items = df, isHeaderVisible = TRUE, selectionMode = "none")
    } else {
      p("No effects added for the current sheet.", style="padding-left: 10px;")
    }
  })
  
  # --- Sheet and Download Logic ---
  observeEvent(input$add_sheet_btn, {
    if (length(current_effects()) == 0) {
      showNotification("Cannot save an empty sheet.", type = "warning", duration = 5)
      return()
    }
    sheet_name <- isolate(input$sheet_name)
    req(nchar(sheet_name) > 0)
    
    current_sheets <- isolate(all_sheets())
    current_sheets[[sheet_name]] <- live_df()
    all_sheets(current_sheets)
    
    removeUI(selector = "#effect_cards_container > div", multiple = TRUE)
    current_effects(list())
    updateTextField.shinyInput(session, "sheet_name", value = paste0("Sheet", length(current_sheets) + 1))
    showNotification(paste("Sheet '", sheet_name, "' saved!", sep=""), type = "success", duration = 5)
  })
  
  output$global_download_btn <- downloadHandler(
    filename = function() {
      paste0("Specification_File-", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      # --- Create final list of sheets to export ---
      final_sheets <- isolate(all_sheets())
      if (length(current_effects()) > 0) {
        final_sheets[[isolate(input$sheet_name)]] <- live_df()
      }
      req(length(final_sheets) > 0, cancelOutput = TRUE)
      
      # --- FIX: Generate QC sheets and add them to the list ---
      sheets_with_qc <- list()
      for (sheet_name in names(final_sheets)) {
        # 1. Add the original specification sheet
        sheets_with_qc[[sheet_name]] <- final_sheets[[sheet_name]]
        
        # 2. Create and add the corresponding QC checklist sheet
        spec_df <- final_sheets[[sheet_name]]
        if ("Effect" %in% names(spec_df) && nrow(spec_df) > 0) {
          qc_df <- data.frame(
            Endpoint = spec_df$Effect,
            `QC Axis Name` = "",
            `QC Value` = "",
            `QC Confidence Intervals` = "",
            `QC Effect Measure Type` = "",
            `QC Axis Directions` = "",
            check.names = FALSE # Important for column names with spaces
          )
          qc_sheet_name <- paste0(sheet_name, "_QC")
          sheets_with_qc[[qc_sheet_name]] <- qc_df
        }
      }
      
      # Write the final combined list to the Excel file
      writexl::write_xlsx(sheets_with_qc, path = file)
    }
  )
}

shinyApp(ui = ui, server = server)
