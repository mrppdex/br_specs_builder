# R Shiny App for Excel Data Validation
#
# INSTRUCTIONS:
# 1. Make sure you have the following R packages installed:
#    install.packages(c("shiny", "shinyjs", "DT", "readxl", "writexl", "dplyr", "purrr", "shinythemes", "shinycssloaders"))
# 2. Save this code as 'app.R' in a new folder.
# 3. Run the app by opening R or RStudio, setting the working directory to that folder, and running: shiny::runApp()

library(shiny)
library(shinyjs)
library(DT)
library(readxl)
library(writexl)
library(dplyr)
library(purrr)
library(shinythemes)
library(shinycssloaders)

# ==============================================================================
# Helper Functions
# ==============================================================================

#' Validate a single value based on a specification rule.
#'
#' @param value The value to validate.
#' @param rule A single row from the specifications data frame.
#' @return A list with `is_valid` (boolean) and `message` (string).
validate_value <- function(value, rule) {
  # Treat empty/NA values as valid by default
  if (is.na(value) || value == "") {
    return(list(is_valid = TRUE, message = NA_character_))
  }
  
  # Coerce value to the correct type for validation
  value_coerced <- tryCatch({
    switch(rule$Type,
           "Numeric" = as.numeric(value),
           "Categorical" = as.character(value),
           "File Path" = as.character(value),
           "Text" = as.character(value)
    )
  }, warning = function(w) { NULL }, error = function(e) { NULL })
  
  
  # Validation logic based on type
  is_valid <- switch(
    rule$Type,
    "Text" = {
      is.character(value_coerced)
    },
    "Numeric" = {
      !is.na(value_coerced) && is.numeric(value_coerced)
    },
    "Categorical" = {
      allowed_values <- strsplit(rule$Values, ",")[[1]] %>% trimws()
      value_coerced %in% allowed_values
    },
    "File Path" = {
      path_val <- trimws(value_coerced)
      is.character(path_val) && (file.exists(path_val) || dir.exists(path_val))
    },
    # Default case
    FALSE 
  )
  
  # Generate message for invalid cells
  message <- if (!is_valid) {
    switch(
      rule$Type,
      "Text" = "Invalid text format.",
      "Numeric" = "Value must be a number.",
      "Categorical" = paste0("Value must be one of: ", rule$Values),
      "File Path" = "File or directory path does not exist on the server."
    )
  } else {
    NA_character_
  }
  
  return(list(is_valid = is_valid, message = message))
}


# ==============================================================================
# UI Definition
# ==============================================================================

ui <- fluidPage(
  theme = shinytheme("spacelab"),
  useShinyjs(), 
  
  titlePanel("Excel Data Quality and Integrity Checker"),
  
  sidebarLayout(
    sidebarPanel(
      width = 3,
      h4("Step 1: Define Validation Specs"),
      helpText("Create specs below or upload a CSV file."),
      
      wellPanel(
          textInput("spec_col_name", "Column Name"),
          selectInput("spec_col_type", "Data Type", 
                      choices = c("Text", "Numeric", "Categorical", "File Path")),
          conditionalPanel(
              condition = "input.spec_col_type == 'Categorical'",
              textAreaInput("spec_col_values", "Allowed Values (comma-separated)")
          ),
          actionButton("add_spec", "Add Spec Rule", icon = icon("plus"))
      ),
      
      h5("Manage Spec Files"),
      fileInput("upload_specs", "Upload Spec CSV", accept = ".csv"),
      downloadButton("download_specs", "Download Specs as CSV"),
      hr(),
      
      h4("Step 2: Upload & Validate Data"),
      fileInput("upload_excel", "Upload Excel File", accept = c(".xlsx")),
      uiOutput("sheet_selector_ui"),
      actionButton("validate_btn", "Validate Selected Sheets", icon = icon("check"), class = "btn-primary"),
      hr(),
      
      h4("User Guide"),
      helpText(
        "1. Add validation rules manually or upload a spec file.",
        "2. Upload the Excel file you want to validate.",
        "3. Select the sheets to validate from the list.",
        "4. Click 'Validate' to see results.",
        "5. View results for each sheet in its own tab. View all issues in 'Error Summary'.",
        "6. Hover over a red cell for an error message.",
        "7. Double-click to edit and correct data. The table will re-validate automatically.",
        "8. Use the 'Export' button on each tab to save the corrected data."
      )
    ),
    
    mainPanel(
      width = 9,
      h4("Specification Rules"),
      withSpinner(DTOutput("spec_table")),
      hr(),
      # FIX: The main content area is now a dynamic UI output
      # This prevents browser rendering bugs from nested tabsets
      uiOutput("main_tabs_ui")
    )
  )
)

# ==============================================================================
# Server Logic
# ==============================================================================

server <- function(input, output, session) {
  
  rv <- reactiveValues(
    specs = tibble(Name = character(), Type = character(), Values = character()),
    uploaded_excel_data = list(),
    validation_results = list(),
    error_summary = tibble()
  )
  
  # --- Spec Management Logic ---
  observeEvent(input$add_spec, {
    req(input$spec_col_name, input$spec_col_type)
    
    if(input$spec_col_name %in% rv$specs$Name){
        showNotification("A rule for this column name already exists.", type = "warning")
        return()
    }

    new_rule <- tibble(
      Name = input$spec_col_name,
      Type = input$spec_col_type,
      Values = if (input$spec_col_type == "Categorical") input$spec_col_values else NA_character_
    )
    rv$specs <- bind_rows(rv$specs, new_rule)
    
    updateTextInput(session, "spec_col_name", value = "")
    updateTextAreaInput(session, "spec_col_values", value = "")
  })
  
  observeEvent(input$upload_specs, {
    req(input$upload_specs)
    df <- try(read.csv(input$upload_specs$datapath, stringsAsFactors = FALSE, check.names = FALSE))
    if(inherits(df, "try-error")){
        showNotification("Failed to read the spec file. Please ensure it's a valid CSV.", type = "error")
        return()
    }
    if(!all(c("Name", "Type", "Values") %in% colnames(df))){
        showNotification("Spec file must contain 'Name', 'Type', and 'Values' columns.", type = "error")
        rv$specs <- tibble(Name = character(), Type = character(), Values = character())
    } else {
        rv$specs <- as_tibble(df)
        showNotification("Specifications loaded successfully.", type = "message")
    }
  })
  
  output$spec_table <- renderDT({
    datatable(rv$specs, 
              editable = TRUE,
              options = list(pageLength = 5, dom = 'tip'),
              rownames = FALSE
    )
  })

  observeEvent(input$spec_table_cell_edit, {
    info <- input$spec_table_cell_edit
    rv$specs <- editData(rv$specs, info, "spec_table")
  })
  
  output$download_specs <- downloadHandler(
    filename = function() { paste0("data-specs-", Sys.Date(), ".csv") },
    content = function(file) { write.csv(rv$specs, file, row.names = FALSE) }
  )
  
  # --- Data Upload Logic ---
  observeEvent(input$upload_excel, {
    req(input$upload_excel)
    path <- input$upload_excel$datapath
    
    tryCatch({
      sheet_names <- excel_sheets(path)
      rv$uploaded_excel_data <- set_names(map(sheet_names, ~read_excel(path, sheet = .x, col_types = "text")), sheet_names)
      
      output$sheet_selector_ui <- renderUI({
        checkboxGroupInput("selected_sheets", "Select Sheets to Validate:",
                           choices = sheet_names, selected = sheet_names)
      })
      showNotification(paste("Excel file loaded with", length(sheet_names), "sheets."), type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error reading Excel file:", e$message), type = "error")
      output$sheet_selector_ui <- renderUI({ helpText("Could not read the uploaded file.") })
    })
  })
  
  # --- Core Validation Function ---
  run_validation <- function() {
    req(input$selected_sheets, nrow(rv$specs) > 0, length(rv$uploaded_excel_data) > 0)
    
    selected <- input$selected_sheets
    specs <- rv$specs
    
    withProgress(message = 'Validating data...', value = 0, {
      
      results <- map(selected, function(sheet_name) {
        if (!sheet_name %in% names(rv$uploaded_excel_data)) return(NULL)
          
        incProgress(1/length(selected), detail = paste("Processing sheet:", sheet_name))
        data_sheet <- rv$uploaded_excel_data[[sheet_name]]
        
        spec_cols <- specs$Name
        data_cols <- colnames(data_sheet)
        missing_cols <- setdiff(spec_cols, data_cols)
        extra_cols <- setdiff(data_cols, spec_cols)
        
        validation_matrix <- matrix(NA_character_, nrow = nrow(data_sheet), ncol = ncol(data_sheet))
        colnames(validation_matrix) <- data_cols
        
        for (spec_rule_row in 1:nrow(specs)) {
          rule <- specs[spec_rule_row, ]
          col_name <- rule$Name
          
          if (col_name %in% data_cols) {
            col_idx <- which(colnames(data_sheet) == col_name)
            for (row_idx in 1:nrow(data_sheet)) {
              value <- data_sheet[[col_name]][row_idx]
              validation_output <- validate_value(value, rule)
              if (!validation_output$is_valid) {
                validation_matrix[row_idx, col_idx] <- validation_output$message
              }
            }
          }
        }
        
        return(list(
          data = data_sheet, validation_matrix = validation_matrix,
          missing_cols = missing_cols, extra_cols = extra_cols
        ))
      })
    }) 
    
    results <- results[!sapply(results, is.null)]
    rv$validation_results <- set_names(results, selected[selected %in% names(rv$uploaded_excel_data)])

    error_summary_df <- imap_dfr(rv$validation_results, ~{
        validation_matrix <- .x$validation_matrix
        data_sheet <- .x$data
        sheet_name <- .y

        error_indices <- which(!is.na(validation_matrix), arr.ind = TRUE)
        
        if(nrow(error_indices) > 0) {
          map_dfr(1:nrow(error_indices), function(i) {
            row_idx <- error_indices[i, 1]
            col_idx <- error_indices[i, 2]
            tibble(
              Sheet = sheet_name,
              Row = row_idx,
              Column = colnames(data_sheet)[col_idx],
              Value = as.character(data_sheet[row_idx, col_idx]),
              Reason = validation_matrix[row_idx, col_idx]
            )
          })
        } else {
          tibble()
        }
    })
    rv$error_summary <- error_summary_df

    if(!is.null(shiny::getDefaultReactiveDomain())) {
        showNotification("Validation complete!", type = "message")
    }
  }

  # --- Event Triggers for Validation ---
  observeEvent(input$validate_btn, {
    run_validation()
  })
  
  # --- Render UI Elements ---
  
  # FIX: This is the new main UI output that builds the entire tabset
  output$main_tabs_ui <- renderUI({
    if (length(rv$validation_results) == 0) {
      return(helpText("Validation results will appear here after you upload a file and click 'Validate'."))
    }
    
    # Create a list of tabs for each sheet
    sheet_tabs <- imap(rv$validation_results, ~{
      sheet_name <- .y
      tabPanel(
        title = sheet_name,
        if (length(.x$missing_cols) > 0) {
          div(style="color:orange; margin-bottom:10px;", paste("Warning: MISSING columns:", paste(.x$missing_cols, collapse=", ")))
        },
        if (length(.x$extra_cols) > 0) {
          div(style="color:#888; margin-bottom:10px;", paste("Info: Extra columns not in specs:", paste(.x$extra_cols, collapse=", ")))
        },
        withSpinner(DTOutput(paste0("table_", sheet_name))),
        downloadButton(paste0("download_", sheet_name), "Export Corrected Sheet as XLSX")
      )
    })
    
    # Create the error summary tab
    summary_tab <- list(tabPanel(
      "Error Summary",
      h4("Consolidated List of Validation Errors"),
      withSpinner(DTOutput("error_summary_table")),
      br(),
      uiOutput("download_errors_ui")
    ))
    
    # Combine sheet tabs and summary tab and create the tabset panel
    all_tabs <- c(sheet_tabs, summary_tab)
    do.call(tabsetPanel, c(id="main_tabset", unname(all_tabs)))
  })
  
  output$error_summary_table <- renderDT({
    if (nrow(rv$error_summary) == 0 && length(rv$validation_results) > 0) {
      return(datatable(data.frame(Message = "No validation errors found across all checked sheets."), 
                       rownames = FALSE, options = list(dom = 't')))
    }
    req(nrow(rv$error_summary) > 0)
    datatable(rv$error_summary,
              rownames = FALSE,
              filter = 'top',
              extensions = 'Buttons',
              options = list(pageLength = 10, scrollX = TRUE, dom = 'Bfrtip', buttons = c('copy', 'csv', 'excel')))
  })

  output$download_errors_ui <- renderUI({
      req(nrow(rv$error_summary) > 0)
      downloadButton("download_error_summary_btn", "Download Full Error Summary")
  })
  
  output$download_error_summary_btn <- downloadHandler(
      filename = function() { paste0("validation-error-summary-", Sys.Date(), ".csv") },
      content = function(file) { write.csv(rv$error_summary, file, row.names = FALSE) }
  )
  
  # --- Dynamic Observers and Outputs for Each Sheet ---
  observe({
    req(length(rv$validation_results) > 0)
    
    walk(names(rv$validation_results), function(sheet_name) {
      
      # Render the DataTable for the sheet
      output[[paste0("table_", sheet_name)]] <- renderDT({
        res <- rv$validation_results[[sheet_name]]
        
        datatable(
          res$data,
          editable = list(target = 'cell'),
          rownames = FALSE,
          options = list(
            pageLength = 10, scrollX = TRUE,
            rowCallback = JS(
              "function(row, data, index) {",
              "  var validationMatrix = ", jsonlite::toJSON(res$validation_matrix, na = "null"), ";",
              # DT's 'index' is 0-based, matching the JS array index.
              "  for (var j=0; j < data.length; j++) {",
              "    if (validationMatrix[index][j] !== null) {",
              "      var cell = $(row).find('td').eq(j);",
              "      cell.attr('title', validationMatrix[index][j]);",
              # Apply CSS styling directly to the invalid cell
              "      cell.css('background-color', 'rgba(255, 135, 135, 0.7)');",
              "    }",
              "  }",
              "}"
            )
          )
        )
      })
      
      # Observer for cell edits
      observeEvent(input[[paste0("table_", sheet_name, "_cell_edit")]], {
        info <- input[[paste0("table_", sheet_name, "_cell_edit")]]
        
        # Update data in reactive value
        rv$uploaded_excel_data[[sheet_name]] <- editData(
          rv$uploaded_excel_data[[sheet_name]], info, rownames = FALSE
        )
        
        # Re-run the full validation after an edit
        run_validation()
      })
      
      # Download handler for the sheet
      output[[paste0("download_", sheet_name)]] <- downloadHandler(
        filename = function() { paste0("corrected-", sheet_name, "-", Sys.Date(), ".xlsx") },
        content = function(file) { write_xlsx(setNames(list(rv$uploaded_excel_data[[sheet_name]]), sheet_name), file) }
      )
    })
  })
}

# Run the application
shinyApp(ui = ui, server = server)
