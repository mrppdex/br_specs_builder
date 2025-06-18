# R Shiny App for Excel Data Validation
#
# INSTRUCTIONS:
# 1. Make sure you have the following R packages installed:
#    install.packages(c("shiny", "shinyjs", "DT", "dplyr", "purrr", "haven", "rlang", "bslib", "openxlsx"))
# 2. Save this code as 'app.R' in a new folder.
# 3. Run the app by opening R or RStudio, setting the working directory to that folder, and running: shiny::runApp()

library(shiny)
library(shinyjs)
library(DT)
library(dplyr)
library(purrr)
library(haven) 
library(rlang) 
library(bslib)      # For modern UI elements
library(openxlsx)   # For all Excel operations

# ==============================================================================
# Helper Functions
# ==============================================================================

#' Validate a single value based on a standard specification rule.
#'
#' @param value The value to validate.
#' @param rule A single row from the specifications data frame.
#' @return A list with `is_valid` (boolean) and `message` (string).
validate_value <- function(value, rule) {
  # FIX: Categorical values are now required and cannot be empty.
  if (rule$Type == "Categorical") {
    if (is.na(value) || value == "") {
      return(list(is_valid = FALSE, message = "Value is required and cannot be empty."))
    }
  } else {
    # For all other types, empty values are considered valid by default.
    if (is.na(value) || value == "") {
      return(list(is_valid = TRUE, message = NA_character_))
    }
  }
  
  value_coerced <- tryCatch({
    switch(rule$Type,
           "Numeric" = as.numeric(value),
           "Categorical" = as.character(value),
           "File Path" = as.character(value),
           "Text" = as.character(value))
  }, warning = function(w) { NULL }, error = function(e) { NULL })
  
  is_valid <- switch(
    rule$Type,
    "Text" = is.character(value_coerced),
    "Numeric" = !is.na(value_coerced) && is.numeric(value_coerced),
    "Categorical" = {
      allowed_values <- strsplit(rule$Values, ",")[[1]] %>% trimws()
      value_coerced %in% allowed_values
    },
    "File Path" = {
      path_val <- trimws(value_coerced)
      is.character(path_val) && (file.exists(path_val) || dir.exists(path_val))
    },
    FALSE 
  )
  
  message <- if (!is_valid) {
    switch(
      rule$Type,
      "Text" = "Invalid text format.",
      "Numeric" = "Value must be a number.",
      "Categorical" = paste0("Value must be one of: ", rule$Values),
      "File Path" = "File or directory path does not exist."
    )
  } else {
    NA_character_
  }
  
  return(list(is_valid = is_valid, message = message))
}

#' Validate ARDS uniqueness based on a file path and filter string.
#'
#' @param ards_path Path to the ARDS dataset (.sas7bdat, .csv, etc.).
#' @param filter_string An R expression string for filtering the data.
#'@param unique_cols_str A comma-separated string of column names for the uniqueness check.
#' @return A list with `is_valid` (boolean) and `message` (string).
validate_ards <- function(ards_path, filter_string, unique_cols_str) {
  if (is.na(ards_path) || ards_path == "" || is.na(filter_string) || filter_string == "") {
    return(list(is_valid = TRUE, message = NA_character_))
  }
  
  if (!file.exists(ards_path)) {
    return(list(is_valid = FALSE, message = paste("ARDS file not found at:", ards_path)))
  }
  
  ards_data <- tryCatch({
    ext <- tools::file_ext(ards_path)
    if (ext == "sas7bdat") haven::read_sas(ards_path)
    else if (ext == "csv") read.csv(ards_path, stringsAsFactors = FALSE)
    else stop(paste("Unsupported file type:", ext))
  }, error = function(e) {
    return(list(is_valid = FALSE, message = paste("Error reading ARDS file:", e$message)))
  })
  
  if (!is.data.frame(ards_data)) return(ards_data) 

  filtered_data <- tryCatch({
    filter_expr <- rlang::parse_expr(filter_string)
    ards_data %>% filter(!!filter_expr)
  }, error = function(e) {
    return(list(is_valid = FALSE, message = paste("Invalid filter syntax:", e$message)))
  })
  
  if (!is.data.frame(filtered_data)) return(filtered_data)
  
  total_rows <- nrow(filtered_data)
  if (total_rows == 0) {
    return(list(is_valid = FALSE, message = "Validation failed: Filter condition resulted in zero rows."))
  }
  
  unique_cols <- strsplit(unique_cols_str, ",")[[1]] %>% trimws()
  if (!all(unique_cols %in% colnames(filtered_data))) {
    missing <- paste(setdiff(unique_cols, colnames(filtered_data)), collapse=", ")
    return(list(is_valid = FALSE, message = paste("Uniqueness columns not found in ARDS data:", missing)))
  }
  
  distinct_rows <- filtered_data %>% 
    distinct(across(all_of(unique_cols))) %>% 
    nrow()
  
  if (total_rows != distinct_rows) {
    msg <- paste0("Uniqueness check failed on columns (", unique_cols_str, "). Filtered data has ", total_rows, " rows, but only ", distinct_rows, " unique combinations.")
    return(list(is_valid = FALSE, message = msg))
  }
  
  return(list(is_valid = TRUE, message = NA_character_))
}


# ==============================================================================
# UI Definition
# ==============================================================================

ui <- page_sidebar(
  theme = bs_theme(version = 5, bootswatch = "spacelab"),
  title = "Specs Quality and Integrity Checker",
  
  sidebar = sidebar(
    accordion(
      open = TRUE,
      accordion_panel(
        "Step 1: Specifications",
        textInput("spec_col_name", "Column Name / Rule Name"),
        selectInput("spec_col_type", "Data Type", 
                    choices = c("Text", "Numeric", "Categorical", "File Path", "ARDS")),
        conditionalPanel(
            condition = "input.spec_col_type == 'Categorical'",
            textAreaInput("spec_col_values", "Allowed Values (comma-separated)")
        ),
        conditionalPanel(
            condition = "input.spec_col_type == 'ARDS'",
            textInput("ards_path_col", "Name of Path Column", placeholder = "e.g., PathToARDS"),
            textInput("ards_filter_col", "Name of Filter Column", placeholder = "e.g., FilterCondition"),
            textInput("ards_unique_cols", "Columns for Uniqueness Check", value = "RESULTTYPE", placeholder = "e.g., RESULTTYPE, RESULT")
        ),
        actionButton("add_spec", "Add Spec Rule", icon = icon("plus")),
        hr(),
        fileInput("upload_specs", "Upload Spec CSV", accept = ".csv"),
        downloadButton("download_specs", "Download Specs as CSV")
      ),
      accordion_panel(
        "Step 2: Upload & Validate",
        fileInput("upload_excel", "Upload Excel File", accept = c(".xlsx")),
        uiOutput("sheet_selector_ui"),
        actionButton("validate_btn", "Validate Selected Sheets", icon = icon("check"), class = "btn-primary"),
        hr(),
        uiOutput("download_all_ui")
      )
    )
  ),
  
  card(
    card_header("Specification Rules"),
    card_body(DTOutput("spec_table"))
  ),
  card(
    card_header("Validation Results"),
    card_body(uiOutput("main_tabs_ui"))
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
    
    spec_values <- NA_character_
    if (input$spec_col_type == "Categorical") {
      spec_values <- input$spec_col_values
    } else if (input$spec_col_type == "ARDS") {
      path_col <- if (is.null(input$ards_path_col)) "" else input$ards_path_col
      filter_col <- if (is.null(input$ards_filter_col)) "" else input$ards_filter_col
      unique_cols <- if (is.null(input$ards_unique_cols)) "RESULTTYPE" else input$ards_unique_cols
      spec_values <- paste0("path_col=", path_col, ";filter_col=", filter_col, ";unique_cols=", unique_cols)
    }

    new_rule <- tibble(
      Name = input$spec_col_name,
      Type = input$spec_col_type,
      Values = spec_values
    )
    rv$specs <- bind_rows(rv$specs, new_rule)
    
    updateTextInput(session, "spec_col_name", value = "")
    updateTextAreaInput(session, "spec_col_values", value = "")
    updateTextInput(session, "ards_path_col", value = "")
    updateTextInput(session, "ards_filter_col", value = "")
    updateTextInput(session, "ards_unique_cols", value = "RESULTTYPE")
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
    datatable(rv$specs, editable = TRUE, options = list(pageLength = 5, dom = 'tip'), rownames = FALSE)
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
      sheet_names <- openxlsx::getSheetNames(path)
      rv$uploaded_excel_data <- set_names(map(sheet_names, ~openxlsx::read.xlsx(path, sheet = .x)), sheet_names)
      
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
        
        validation_matrix <- matrix(NA_character_, nrow = nrow(data_sheet), ncol = ncol(data_sheet))
        colnames(validation_matrix) <- colnames(data_sheet)
        
        append_error <- function(existing_msg, new_msg) {
          if (is.na(existing_msg)) return(new_msg)
          else return(paste(existing_msg, new_msg, sep = " | "))
        }

        standard_specs <- specs %>% filter(Type != "ARDS")
        ards_specs <- specs %>% filter(Type == "ARDS")

        if(nrow(standard_specs) > 0) {
          for (spec_rule_row in 1:nrow(standard_specs)) {
            rule <- standard_specs[spec_rule_row, ]
            col_name <- rule$Name
            if (col_name %in% colnames(data_sheet)) {
              col_idx <- which(colnames(data_sheet) == col_name)
              for (row_idx in 1:nrow(data_sheet)) {
                value <- data_sheet[[col_name]][row_idx]
                validation_output <- validate_value(value, rule)
                if (!validation_output$is_valid) {
                  validation_matrix[row_idx, col_idx] <- append_error(validation_matrix[row_idx, col_idx], validation_output$message)
                }
              }
            }
          }
        }
        
        if(nrow(ards_specs) > 0) {
          for (spec_rule_row in 1:nrow(ards_specs)) {
            rule <- ards_specs[spec_rule_row, ]
            params <- strsplit(rule$Values, ";")[[1]]
            get_param <- function(p_name) {
              val <- params[grepl(paste0("^", p_name, "="), params)]
              if(length(val) == 0) return(NA_character_)
              sub(paste0(p_name, "="), "", val)
            }
            path_col_name <- get_param("path_col")
            filter_col_name <- get_param("filter_col")
            unique_cols_str <- get_param("unique_cols")
            if(is.na(unique_cols_str)) unique_cols_str <- "RESULTTYPE"
            if (is.na(path_col_name) || is.na(filter_col_name) || !all(c(path_col_name, filter_col_name) %in% colnames(data_sheet))) next
            
            path_col_idx <- which(colnames(data_sheet) == path_col_name)
            for (row_idx in 1:nrow(data_sheet)) {
              ards_path <- data_sheet[[path_col_name]][row_idx]
              filter_str <- data_sheet[[filter_col_name]][row_idx]
              validation_output <- validate_ards(ards_path, filter_str, unique_cols_str)
              if (!validation_output$is_valid) {
                 validation_matrix[row_idx, path_col_idx] <- append_error(validation_matrix[row_idx, path_col_idx], validation_output$message)
              }
            }
          }
        }
        
        standard_col_names <- standard_specs$Name
        ards_dependent_cols <- if(nrow(ards_specs) > 0) {
            purrr::map(ards_specs$Values, function(v) {
                p <- strsplit(v, ";")[[1]]
                c(sub("path_col=", "", p[grepl("path_col=", p)]), sub("filter_col=", "", p[grepl("filter_col=", p)]))
            }) %>% unlist() %>% unique()
        } else { c() }
        all_expected_cols <- unique(c(standard_col_names, ards_dependent_cols))

        return(list(data = data_sheet, validation_matrix = validation_matrix,
          missing_cols = setdiff(all_expected_cols, colnames(data_sheet)), 
          extra_cols = setdiff(colnames(data_sheet), all_expected_cols)))
      })
    }) 
    
    results <- results[!sapply(results, is.null)]
    rv$validation_results <- set_names(results, selected[selected %in% names(rv$uploaded_excel_data)])

    error_summary_df <- imap_dfr(rv$validation_results, ~{
        validation_matrix <- .x$validation_matrix
        error_indices <- which(!is.na(validation_matrix), arr.ind = TRUE)
        if(nrow(error_indices) > 0) {
          map_dfr(1:nrow(error_indices), function(i) {
            row_idx <- error_indices[i, "row"]; col_idx <- error_indices[i, "col"]
            tibble(Sheet = .y, Row = row_idx, Column = colnames(.x$data)[col_idx],
                   Value = as.character(.x$data[row_idx, col_idx]), Reason = validation_matrix[row_idx, col_idx])
          })
        } else { tibble() }
    })
    rv$error_summary <- error_summary_df

    if(!is.null(shiny::getDefaultReactiveDomain())) { showNotification("Validation complete!", type = "message") }
  }

  # --- Event Triggers for Validation ---
  observeEvent(input$validate_btn, { run_validation() })
  
  # --- Render UI Elements ---
  output$main_tabs_ui <- renderUI({
    if (length(rv$validation_results) == 0) {
      return(helpText("Validation results will appear here after you upload a file and click 'Validate'."))
    }
    
    sheet_tabs <- imap(rv$validation_results, ~{
      tabPanel(title = .y,
        if (length(.x$missing_cols) > 0) {
          div(style="color:orange; margin-bottom:10px;", paste("Warning: MISSING columns:", paste(.x$missing_cols, collapse=", ")))
        },
        if (length(.x$extra_cols) > 0) {
          div(style="color:#888; margin-bottom:10px;", paste("Info: Extra columns not in specs:", paste(.x$extra_cols, collapse=", ")))
        },
        DTOutput(paste0("table_", .y))
      )
    })
    
    summary_tab <- list(tabPanel("Error Summary",
      h4("Consolidated List of Validation Errors"), DTOutput("error_summary_table"), br(), uiOutput("download_errors_ui")))
    
    do.call(tabsetPanel, c(id="main_tabset", unname(c(sheet_tabs, summary_tab))))
  })
  
  output$error_summary_table <- renderDT({
    if (nrow(rv$error_summary) == 0 && length(rv$validation_results) > 0) {
      return(datatable(data.frame(Message = "No validation errors found across all checked sheets."), rownames = FALSE, options = list(dom = 't')))
    }
    req(nrow(rv$error_summary) > 0)
    datatable(rv$error_summary, rownames = FALSE, filter = 'top',
              options = list(pageLength = 10, scrollX = TRUE, dom = 'frtip'))
  })

  output$download_errors_ui <- renderUI({
    req(nrow(rv$error_summary) > 0)
    downloadButton("download_error_summary_btn", "Download Full Error Report")
  })
  
  # FIX: Updated download handler for advanced Excel formatting
  output$download_error_summary_btn <- downloadHandler(
    filename = function() { paste0("validation-error-summary-", Sys.Date(), ".xlsx") },
    content = function(file) {
      req(nrow(rv$error_summary) > 0)
      
      # Sort data
      sorted_summary <- rv$error_summary %>%
        arrange(Sheet, as.numeric(Row), Column)
        
      # Create a new workbook
      wb <- createWorkbook()
      addWorksheet(wb, "Error Summary")
      
      # Add and merge the main title
      writeData(wb, "Error Summary", "Specs Quality and Integrity Checker", startCol = 1, startRow = 1)
      mergeCells(wb, "Error Summary", cols = 1:ncol(sorted_summary), rows = 1)
      
      # Style for the title (bold, centered)
      titleStyle <- createStyle(fontSize = 14, textDecoration = "bold", halign = "center")
      addStyle(wb, "Error Summary", style = titleStyle, rows = 1, cols = 1)
      
      # Write the data frame starting from row 3
      writeData(wb, "Error Summary", sorted_summary, startRow = 3)
      
      # Style for the header (bold)
      headerStyle <- createStyle(textDecoration = "bold")
      addStyle(wb, "Error Summary", style = headerStyle, rows = 3, cols = 1:ncol(sorted_summary), gridExpand = TRUE)
      
      # Set column widths to auto
      setColWidths(wb, "Error Summary", cols = 1:ncol(sorted_summary), widths = "auto")
      
      # Save the workbook
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  # --- UI for Single Download Button ---
  output$download_all_ui <- renderUI({
    req(length(rv$validation_results) > 0)
    downloadButton("download_all_sheets_btn", "Export All Corrected Sheets", icon = icon("download"))
  })

  # --- Download Handler for All Sheets ---
  output$download_all_sheets_btn <- downloadHandler(
    filename = function() {
        paste0("corrected-data-", Sys.Date(), ".xlsx")
    },
    content = function(file) {
        sheets_to_write <- rv$uploaded_excel_data[names(rv$validation_results)]
        openxlsx::write.xlsx(sheets_to_write, file)
    }
  )

  # --- Dynamic Observers and Outputs for Each Sheet ---
  observe({
    req(length(rv$validation_results) > 0)
    walk(names(rv$validation_results), function(sheet_name) {
      
      output[[paste0("table_", sheet_name)]] <- renderDT({
        res <- rv$validation_results[[sheet_name]]
        datatable(res$data, editable = list(target = 'cell'), rownames = FALSE,
          options = list(pageLength = 10, scrollX = TRUE,
            rowCallback = JS(
              "function(row, data, index) {",
              "  var validationMatrix = ", jsonlite::toJSON(res$validation_matrix, na = "null"), ";",
              "  for (var j=0; j < data.length; j++) {",
              "    if (validationMatrix[index] && validationMatrix[index][j] !== null) {",
              "      var cell = $(row).find('td').eq(j);",
              "      cell.attr('title', validationMatrix[index][j]);",
              "      cell.css('background-color', 'rgba(255, 135, 135, 0.7)');",
              "    }",
              "  }",
              "}"
            )
          )
        )
      })
      
      observeEvent(input[[paste0("table_", sheet_name, "_cell_edit")]], {
        info <- input[[paste0("table_", sheet_name, "_cell_edit")]]
        rv$uploaded_excel_data[[sheet_name]][info$row, info$col] <- as.character(info$value)
        run_validation()
      })
    })
  })
}

# Run the application
shinyApp(ui = ui, server = server)
