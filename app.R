# R Shiny App for Excel Data Validation
#
# INSTRUCTIONS:
# 1. Make sure you have the following R packages installed:
#    install.packages(c("shiny", "shinyjs", "DT", "dplyr", "purrr", "haven", "rlang", "bslib", "openxlsx", "jsonlite"))
# 2. Save this code as 'app.R' in a new folder.
# 3. Run the app by opening R or RStudio, setting the working directory to that folder, and running: shiny::runApp()

library(shiny)
library(shinyjs)
library(DT)
library(dplyr)
library(purrr)
library(haven) 
library(rlang) 
library(bslib)
library(openxlsx)
library(jsonlite) # For parsing JSON in ARDS filter

# ==============================================================================
# Helper Functions
# ==============================================================================

#' Validate a single value based on a standard specification rule.
#'
#' @param value The value to validate.
#' @param rule A single row from the specifications data frame.
#' @return A list with `is_valid` (boolean) and `message` (string).
validate_value <- function(value, rule) {
  if (rule$Type == "Categorical") {
    if (is.na(value) || value == "") {
      return(list(is_valid = FALSE, message = "Value is required and cannot be empty."))
    }
  } else {
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
  } else { NA_character_ }
  
  return(list(is_valid = is_valid, message = message))
}

#' A placeholder function to extract and process data from an ARDS dataset.
#' This function is designed to be replaced with your specific business logic.
#'
#' @return A list containing `is_valid` (boolean) and a `message`.
#'   - On success, `is_valid` is TRUE and `message` is the resulting dataframe.
#'   - On failure, `is_valid` is FALSE and `message` is an error string.
extract_from_ards_with_query <- function(ards_data, filter_query, measure, difference_measure, difference_lci, difference_uci, cmp_name, ref_column, trt_column) {
  # This is a placeholder for your actual logic.
  # You would use the parameters to query the ards_data.
  
  # Example of a success case:
  # For demonstration, we'll just filter the ARDS data and return the head.
  # In a real scenario, you would perform your complex extraction here.
  df_result <- tryCatch({
    ards_data %>%
      filter(!!rlang::parse_expr(filter_query)) %>%
      head() # Return first few rows as an example
  }, error = function(e) {
    # Return an error if the filter fails inside this function
    return(list(is_valid = FALSE, message = paste("Error during data extraction:", e$message)))
  })
  
  # If the above tryCatch resulted in an error list, return it
  if (!is.data.frame(df_result)) return(df_result)
  
  # If we get a result, return it as the message
  return(list(is_valid = TRUE, message = df_result))
}

#' Orchestrator for ARDS validation
#'
#' @return A list with `is_valid` (boolean) and `message` (string or dataframe).
validate_ards <- function(ards_path, json_string) {
  if (is.na(ards_path) || ards_path == "" || is.na(json_string) || json_string == "") {
    return(list(is_valid = TRUE, message = NA_character_))
  }
  
  if (!file.exists(ards_path)) {
    return(list(is_valid = FALSE, message = paste("ARDS file not found at:", ards_path)))
  }
  
  # Parse JSON from the filter string
  params <- tryCatch({
    fromJSON(json_string)
  }, error = function(e) {
    return(list(is_valid = FALSE, message = paste("Invalid JSON in filter column:", e$message)))
  })
  if (!is.list(params)) return(params) # Return error if parsing failed
  
  # Read the dataset
  ards_data <- tryCatch({
    ext <- tools::file_ext(ards_path)
    if (ext == "sas7bdat") haven::read_sas(ards_path)
    else if (ext == "csv") read.csv(ards_path, stringsAsFactors = FALSE)
    else stop(paste("Unsupported file type:", ext))
  }, error = function(e) {
    return(list(is_valid = FALSE, message = paste("Error reading ARDS file:", e$message)))
  })
  if (!is.data.frame(ards_data)) return(ards_data) 
  
  # Call the new extraction function with parameters from JSON
  # Provide defaults as requested
  result <- extract_from_ards_with_query(
    ards_data = ards_data,
    filter_query = params$filter,
    measure = params$measure %||% NA,
    difference_measure = params$difference_measure %||% NA,
    difference_lci = params$difference_lci %||% NA,
    difference_uci = params$difference_uci %||% NA,
    cmp_name = params$cmp_name %||% "Placebo",
    ref_column = params$ref_column %||% "reftrt",
    trt_column = params$trt_column %||% "trt"
  )
  
  return(result)
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
        # FIX: Simplified ARDS UI
        conditionalPanel(
          condition = "input.spec_col_type == 'ARDS'",
          textInput("ards_path_col", "Name of Path Column", placeholder = "e.g., PathToARDS"),
          textInput("ards_filter_col", "Name of JSON Filter Column", placeholder = "e.g., FilterCondition")
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
        actionButton("validate_btn", "Validate Selected Sheets", icon = icon("check"), class = "btn-primary")
        # FIX: Removed download all button as editing is disabled
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
      # FIX: Removed uniqueness cols from spec
      spec_values <- paste0("path_col=", path_col, ";filter_col=", filter_col)
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
    # FIX: Editing removed
    datatable(rv$specs, options = list(pageLength = 5, dom = 'tip'), rownames = FALSE)
  })
  
  # FIX: Cell edit observer removed
  
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
        
        # Create separate matrices for errors (red highlight) and tooltips (hover text)
        error_matrix <- matrix(NA_character_, nrow = nrow(data_sheet), ncol = ncol(data_sheet))
        tooltip_matrix <- matrix(NA_character_, nrow = nrow(data_sheet), ncol = ncol(data_sheet))
        colnames(error_matrix) <- colnames(tooltip_matrix) <- colnames(data_sheet)
        
        append_msg <- function(existing_msg, new_msg) {
          if (is.na(existing_msg)) return(new_msg) else return(paste(existing_msg, new_msg, sep = " | "))
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
                  error_matrix[row_idx, col_idx] <- append_msg(error_matrix[row_idx, col_idx], validation_output$message)
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
            if (is.na(path_col_name) || is.na(filter_col_name) || !all(c(path_col_name, filter_col_name) %in% colnames(data_sheet))) next
            
            # FIX: Target the filter column for messages and highlighting
            filter_col_idx <- which(colnames(data_sheet) == filter_col_name)
            for (row_idx in 1:nrow(data_sheet)) {
              ards_path <- data_sheet[[path_col_name]][row_idx]
              json_str <- data_sheet[[filter_col_name]][row_idx]
              
              validation_output <- validate_ards(ards_path, json_str)
              
              # Add message to tooltip matrix regardless of validity
              if (is.data.frame(validation_output$message)) {
                # Format successful dataframe result for tooltip
                formatted_df <- paste(capture.output(print(validation_output$message)), collapse = "\n")
                tooltip_matrix[row_idx, filter_col_idx] <- append_msg(tooltip_matrix[row_idx, filter_col_idx], formatted_df)
              } else if (!is.na(validation_output$message)) {
                tooltip_matrix[row_idx, filter_col_idx] <- append_msg(tooltip_matrix[row_idx, filter_col_idx], validation_output$message)
              }
              
              # Add message to error matrix ONLY if invalid
              if (!validation_output$is_valid) {
                error_matrix[row_idx, filter_col_idx] <- append_msg(error_matrix[row_idx, filter_col_idx], validation_output$message)
              }
            }
          }
        }
        
        # Combine all tooltips from errors
        tooltip_matrix[!is.na(error_matrix)] <- error_matrix[!is.na(error_matrix)]
        
        standard_col_names <- standard_specs$Name
        ards_dependent_cols <- if(nrow(ards_specs) > 0) {
          purrr::map(ards_specs$Values, function(v) {
            p <- strsplit(v, ";")[[1]]
            c(sub("path_col=", "", p[grepl("path_col=", p)]), sub("filter_col=", "", p[grepl("filter_col=", p)]))
          }) %>% unlist() %>% unique()
        } else { c() }
        all_expected_cols <- unique(c(standard_col_names, ards_dependent_cols))
        
        return(list(data = data_sheet, error_matrix = error_matrix, tooltip_matrix = tooltip_matrix,
                    missing_cols = setdiff(all_expected_cols, colnames(data_sheet)), 
                    extra_cols = setdiff(colnames(data_sheet), all_expected_cols)))
      })
    }) 
    
    results <- results[!sapply(results, is.null)]
    rv$validation_results <- set_names(results, selected[selected %in% names(rv$uploaded_excel_data)])
    
    error_summary_df <- imap_dfr(rv$validation_results, ~{
      error_matrix <- .x$error_matrix
      error_indices <- which(!is.na(error_matrix), arr.ind = TRUE)
      if(nrow(error_indices) > 0) {
        map_dfr(1:nrow(error_indices), function(i) {
          row_idx <- error_indices[i, "row"]; col_idx <- error_indices[i, "col"]
          tibble(Sheet = .y, Row = row_idx, Column = colnames(.x$data)[col_idx],
                 Value = as.character(.x$data[row_idx, col_idx]), Reason = error_matrix[row_idx, col_idx])
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
  
  output$download_error_summary_btn <- downloadHandler(
    filename = function() { paste0("validation-error-summary-", Sys.Date(), ".xlsx") },
    content = function(file) {
      req(nrow(rv$error_summary) > 0)
      sorted_summary <- rv$error_summary %>% arrange(Sheet, as.numeric(Row), Column)
      wb <- createWorkbook()
      addWorksheet(wb, "Error Summary")
      writeData(wb, "Error Summary", "Specs Quality and Integrity Checker", startCol = 1, startRow = 1)
      mergeCells(wb, "Error Summary", cols = 1:ncol(sorted_summary), rows = 1)
      titleStyle <- createStyle(fontSize = 14, textDecoration = "bold", halign = "center")
      addStyle(wb, "Error Summary", style = titleStyle, rows = 1, cols = 1)
      writeData(wb, "Error Summary", sorted_summary, startRow = 3)
      headerStyle <- createStyle(textDecoration = "bold")
      addStyle(wb, "Error Summary", style = headerStyle, rows = 3, cols = 1:ncol(sorted_summary), gridExpand = TRUE)
      setColWidths(wb, "Error Summary", cols = 1:ncol(sorted_summary), widths = "auto")
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  # --- Dynamic Observers and Outputs for Each Sheet ---
  observe({
    req(length(rv$validation_results) > 0)
    walk(names(rv$validation_results), function(sheet_name) {
      
      output[[paste0("table_", sheet_name)]] <- renderDT({
        res <- rv$validation_results[[sheet_name]]
        datatable(res$data, rownames = FALSE, # FIX: Editing removed
                  options = list(pageLength = 10, scrollX = TRUE,
                                 rowCallback = JS(
                                   "function(row, data, index) {",
                                   "  var errorMatrix = ", jsonlite::toJSON(res$error_matrix, na = "null"), ";",
                                   "  var tooltipMatrix = ", jsonlite::toJSON(res$tooltip_matrix, na = "null"), ";",
                                   "  for (var j=0; j < data.length; j++) {",
                                   "    if (tooltipMatrix[index] && tooltipMatrix[index][j] !== null) {",
                                   "      var cell = $(row).find('td').eq(j);",
                                   "      cell.attr('title', tooltipMatrix[index][j]);",
                                   "    }",
                                   "    if (errorMatrix[index] && errorMatrix[index][j] !== null) {",
                                   "      var cell = $(row).find('td').eq(j);",
                                   "      cell.css('background-color', 'rgba(255, 135, 135, 0.7)');",
                                   "    }",
                                   "  }",
                                   "}"
                                 )
                  )
        )
      })
    })
  })
}

# Run the application
shinyApp(ui = ui, server = server)
