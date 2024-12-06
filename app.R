library(shiny)
library(DT)
library(RSQLite)
library(ggplot2)
library(openxlsx)
library(dplyr)
library(zip)
library(plotly)
library(pdftools)
library(shinyjs)
library(stringr)
library(gridExtra)
library(grid)
library(gt)
library(jpeg)
library(shiny)
library(shinydashboard)
library(DT)
library(plotly)

ui <- dashboardPage(
  dashboardHeader(
    title = div(
      icon("pen-to-square"),
      "EQA Report Review"
    ),
    titleWidth = 450
  ),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Upload Data", tabName = "upload_data", icon = icon("file-upload")),
      menuItem("Results", tabName = "results", icon = icon("table")),
      menuItem("Plots", tabName = "plots", icon = icon("chart-line")),
      menuItem("Download", tabName = "download", icon = icon("download"))
    )
  ),
  dashboardBody(
    tags$head(
      tags$title("EQA Report Review"), # Set browser tab title
      tags$link(rel = "icon", type = "image/png", href = "pencil.png"),
      tags$style(HTML("
        .main-header .logo {
          font-size: 30px;  /* Make the title larger */
          font-weight: bold;
          color: #007bff;
          text-transform: none;
        }
        .main-header .logo i {
          font-size: 36px; /* Adjust icon size */
          margin-right: 10px; /* Space between icon and title */
        }
        .navbar {
          min-height: 80px; /* Ensure the navbar height matches */
        }
        #uploadSuccess {
          position: fixed;
          top: 50%;
          left: 50%;
          transform: translate(-50%, -50%);
          font-size: 48px; /* Adjusted for better visibility */
          text-align: center; /* Center-align text inside the notification */
          z-index: 1050;
        }
      "))
    ),
    tabItems(
      tabItem(tabName = "upload_data", fluidRow(
        box(title = "Upload Data", width = 12, fileInput("pdf_files", "Upload PDF Files", multiple = TRUE, accept = ".pdf"))
      )),
      tabItem(tabName = "results", fluidRow(
        box(title = "Results", width = 12, DT::dataTableOutput("data_table"))
      )),
      tabItem(tabName = "plots", fluidRow(
        box(title = "Select Analyte", width = 12,
            selectInput("analyte_plot", "Select Analyte", choices = NULL)
        ),
        box(title = "Regression Plot", width = 12,
            plotlyOutput("regression_plot", height = "500px"),
            checkboxInput("exclude_outliers_regression", "Exclude outliers for Regression Plot?", value = FALSE)
        ),
        box(title = "Bias Plot", width = 12,
            plotlyOutput("bias_plot", height = "500px"),
            checkboxInput("exclude_outliers_bias", "Exclude outliers for Bias Plot?", value = FALSE)
        ),
        box(title = "Z-Score Plot", width = 12,
            plotlyOutput("z_score_plot", height = "500px"),
            checkboxInput("exclude_outliers_zscore", "Exclude outliers?", value = FALSE)  # Moved here
        ),
        box(title = "Frequency Plot", width = 12,
            plotlyOutput("additional_plot", height = "500px")
        )
      )),
      tabItem(tabName = "download", fluidRow(
        column(width = 6,
               box(title = "Download Excel Report", width = 12,
                   checkboxGroupInput("analyte_excel",
                                      "Select Analytes for Excel Report:",
                                      choices = NULL,
                                      selected = NULL,
                                      inline = TRUE),
                   downloadButton("download_report",
                                  "Download Report (Excel)",
                                  class = "btn-primary")
               )
        ),
        column(width = 6,
               box(title = "Download PDF Report", width = 12,
                   checkboxGroupInput("analyte_pdf",
                                      "Select Analytes for PDF Report:",
                                      choices = NULL,
                                      selected = NULL,
                                      inline = TRUE),
                   downloadButton("download_summary_report",
                                  "Download Summary Report (PDF)",
                                  class = "btn-primary")
               )
        )
      ))
    )
  )
)

# Define Server Logic
server <- function(input, output, session) {

  extracted_data <- reactiveVal(data.frame())  # Initialize the reactive variable

  # Database connection setup - Initialize outside observeEvent
  db_connection <- dbConnect(RSQLite::SQLite(), "results_monitoring.db")

  # Database table creation (only if it doesn't exist)
  if (!dbExistsTable(db_connection, "extracted_data")) {
    dbExecute(db_connection, "
      CREATE TABLE IF NOT EXISTS extracted_data (
        Method_Set TEXT,
        Sample TEXT,
        Analyte TEXT,
        Your_Method TEXT,
        Your_Result TEXT,
        N_Method TEXT,
        N_All_Labs TEXT,
        Mean_Method TEXT,
        Mean_All_Labs TEXT,
        Median_Method TEXT,
        Median_All_Labs TEXT,
        SD_Method TEXT,
        SD_All_Labs TEXT,
        Z_score_Method REAL,
        Z_score_All_Labs REAL,
        Bias_Method REAL,
        Bias_All_Labs REAL
      )
    ")
  }

  # Define a safe conversion function
  safe_as_numeric <- function(x) {
    # Attempt to convert x to numeric, and return NA if it can't be converted
    as.numeric(as.character(x))
  }

  # Handle PDF file uploads automatically
  observeEvent(input$pdf_files, {
    req(input$pdf_files)

    # Initialize an empty dataframe to store extracted data
    final_results <- data.frame(
      Method_Set = character(),
      Sample = character(),
      Analyte = character(),
      Your_Method = character(),
      Your_Result = character(),
      N_Method = character(),
      N_All_Labs = character(),
      Mean_Method = character(),
      Mean_All_Labs = character(),
      Median_Method = character(),
      Median_All_Labs = character(),
      SD_Method = character(),
      SD_All_Labs = character(),
      Z_score_Method = numeric(),
      Z_score_All_Labs = numeric(),
      Bias_Method = numeric(),
      Bias_All_Labs = numeric(),
      stringsAsFactors = FALSE
    )

    for (file in input$pdf_files$datapath) {
      results <- pdf_text(file)

      for (page_text in results) {
        # Use safe regex matching
        method_set <- str_trim(str_extract(page_text, "(?i)(?<=Method Set : )[^\\n]+"))
        sample <- str_trim(str_extract(page_text, "(?i)(?<=Sample : )[^\\n]+"))
        analyte <- str_trim(str_extract(page_text, "(?i)(?<=Analyte : )[^\\n]+"))
        your_method <- str_trim(str_extract(page_text, "(?i)(?<=Your Method : )[^\\n]+"))
        your_result <- str_trim(str_extract(page_text, "(?i)(?<=Your Result : )[^\\n]+"))

        # Handle missing values carefully
        n_method_all <- str_match(page_text, "(?i)n\\s*:\\s*(\\d+)\\s+(\\d+)")
        n_method <- ifelse(!is.na(n_method_all[2]), as.character(n_method_all[2]), NA)
        n_all_labs <- ifelse(!is.na(n_method_all[3]), as.character(n_method_all[3]), NA)

        mean_method_all <- str_match(page_text, "(?i)Mean\\s*:\\s*([\\d.]+)\\s+([\\d.]+)")
        mean_method <- ifelse(!is.na(mean_method_all[2]), safe_as_numeric(mean_method_all[2]), NA)
        mean_all_labs <- ifelse(!is.na(mean_method_all[3]), safe_as_numeric(mean_method_all[3]), NA)

        median_method_all <- str_match(page_text, "(?i)Median\\s*:\\s*([\\d.]+)\\s+([\\d.]+)")
        median_method <- ifelse(!is.na(median_method_all[2]), safe_as_numeric(median_method_all[2]), NA)
        median_all_labs <- ifelse(!is.na(median_method_all[3]), safe_as_numeric(median_method_all[3]), NA)

        sd_method_all <- str_match(page_text, "(?i)SD\\s*:\\s*([\\d.]+)\\s+([\\d.]+)")
        sd_method <- ifelse(!is.na(sd_method_all[2]), safe_as_numeric(sd_method_all[2]), NA)
        sd_all_labs <- ifelse(!is.na(sd_method_all[3]), safe_as_numeric(sd_method_all[3]), NA)

        z_score_method <- ifelse(!is.na(your_result) && !is.na(mean_method) && !is.na(sd_method),
                                 round((safe_as_numeric(your_result) - mean_method) / sd_method, 2), NA)

        z_score_all_labs <- ifelse(!is.na(your_result) && !is.na(mean_all_labs) && !is.na(sd_all_labs),
                                   round((safe_as_numeric(your_result) - mean_all_labs) / sd_all_labs, 2), NA)

        bias_method <- ifelse(!is.na(your_result) && !is.na(mean_method) && !is.na(mean_method),
                              round(((mean_method - safe_as_numeric(your_result)) / safe_as_numeric(mean_method)) * 100, 2), NA)

        bias_all_labs <- ifelse(!is.na(your_result) && !is.na(mean_all_labs) && !is.na(mean_all_labs),
                                round(((mean_all_labs - safe_as_numeric(your_result)) / safe_as_numeric(mean_all_labs)) * 100, 2), NA)

        # Add the new row to final results
        new_row <- data.frame(
          Method_Set = method_set,
          Sample = sample,
          Analyte = analyte,
          Your_Method = your_method,
          Your_Result = your_result,
          N_Method = n_method,
          N_All_Labs = n_all_labs,
          Mean_Method = mean_method,
          Mean_All_Labs = mean_all_labs,
          Median_Method = median_method,
          Median_All_Labs = median_all_labs,
          SD_Method = sd_method,
          SD_All_Labs = sd_all_labs,
          Z_score_Method = z_score_method,
          Z_score_All_Labs = z_score_all_labs,
          Bias_Method = bias_method,
          Bias_All_Labs = bias_all_labs,
          stringsAsFactors = FALSE
        )

        # Append the new row to the final results
        final_results <- rbind(final_results, new_row)
      }
    }

    # Update the reactive data with the final results
    extracted_data(final_results)

    # Save data to database
    dbWriteTable(db_connection, "extracted_data", final_results, append = TRUE, row.names = FALSE)

    # Notify user that files were uploaded successfully with added excitement and dynamic features
    showNotification(
      HTML('<i class="fa fa-check-circle" style="color: white;"></i> Files uploaded successfully! 🎉'),
      type = "message",
      duration = 5,
      closeButton = TRUE,
      id = "uploadSuccess",
      session = session
    )
  })

  # Show table of extracted data
  output$data_table <- DT::renderDataTable({
    req(extracted_data())  # Ensure the data is available
    DT::datatable(extracted_data(), options = list(scrollX = TRUE))
  })

  # Populate analyte choices once data is uploaded
  observeEvent(extracted_data(), {
    # Update the analyte selection input based on the data
    analytes <- unique(extracted_data()$Analyte)

    updateSelectInput(session, "analyte_plot",
                      choices = analytes)

    # Update the checkboxes for Excel and PDF reports
    updateCheckboxGroupInput(session, "analyte_excel",
                             choices = analytes,
                             selected = analytes[1])

    updateCheckboxGroupInput(session, "analyte_pdf",
                             choices = analytes,
                             selected = analytes[1])
  })

  # Plots rendering with input validation
  # Bias Plot
  output$bias_plot <- renderPlotly({
    analyte <- input$analyte_plot
    exclude_outliers <- input$exclude_outliers_bias  # Use the bias-specific checkbox

    if (!is.null(analyte) && analyte != "") {
      filtered_data <- extracted_data() %>%
        filter(Analyte == analyte)

      if (exclude_outliers) {
        filtered_data <- filtered_data %>%
          filter(abs(Bias_Method) <= 15)
      }

      # Create a new column to define the labels based on the bias value
      filtered_data <- filtered_data %>%
        mutate(Bias_Label = ifelse(abs(Bias_Method) > 15, "Action Required", "OK"))

      p <- ggplot(filtered_data, aes(x = Sample, y = Bias_Method)) +
        geom_point(aes(color = Bias_Label), shape = 4, size = 3) +
        geom_line(color = "black") +
        geom_hline(yintercept = 0, color = "black", linetype = "dashed") +
        geom_hline(yintercept = c(-10, 10), color = "red", linetype = "longdash") +
        labs(title = analyte,
             x = "Sample",
             y = "Bias (%, Method Group)") +
        scale_color_manual(values = c("Action Required" = "firebrick1", "OK" = "black"),
                           name = "Bias",
                           labels = c("Action Required" = "Bias >15%", "OK" = "Bias <=15%")) +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1))

      ggplotly(p)
    }
  })

  # Regression Plot with Line of Best Fit and R-squared
  output$regression_plot <- renderPlotly({
    analyte <- input$analyte_plot
    exclude_outliers <- input$exclude_outliers_regression  # Use the regression-specific checkbox

    if (!is.null(analyte) && analyte != "") {
      filtered_data <- extracted_data() %>%
        filter(Analyte == analyte)

      # Exclude outliers if the checkbox is checked
      if (exclude_outliers) {
        filtered_data <- filtered_data %>%
          filter(abs(Bias_Method) < 3)
      }

      # Ensure numeric data for regression
      filtered_data$Your_Result <- as.numeric(filtered_data$Your_Result)
      filtered_data$Mean_Method <- as.numeric(filtered_data$Mean_Method)

      # Check for any NA values in the relevant columns
      filtered_data <- filtered_data %>%
        filter(!is.na(Your_Result) & !is.na(Mean_Method))

      # Calculate the linear model and R-squared
      lm_model <- lm(Mean_Method ~ Your_Result, data = filtered_data)
      r_squared <- summary(lm_model)$r.squared

      # Create the regression plot
      p <- ggplot(filtered_data, aes(x = Your_Result, y = Mean_Method)) +
        geom_point(aes(color = Sample), shape = 4, size = 3) +  # Points for each sample
        geom_smooth(method = "lm", se = FALSE, color = "blue") +  # Line of best fit
        labs(
          title = analyte,
          x = "Your Result, umol/L",
          y = "Mean Result (Method Group), umol/L"
        ) +
        annotate(
          "text",
          x = min(filtered_data$Your_Result, na.rm = TRUE),  # Position near the left of the plot
          y = max(filtered_data$Mean_Method, na.rm = TRUE),  # Position near the top of the plot
          label = paste("R² =", round(r_squared, 3)),
          hjust = 1, vjust = 1, size = 3.5, color = "red", fontface = "bold"
        ) +
        theme_minimal() +
        theme(
          axis.title = element_text(size = 10),  # Increase axis title size
          axis.text.x = element_text(size = 10, hjust = 1) # Increase axis text size
        )

      ggplotly(p)
    }
  })


  # Z-Score Plot
  output$z_score_plot <- renderPlotly({
    analyte <- input$analyte_plot
    exclude_outliers <- input$exclude_outliers_zscore  # Get the value from the checkbox

    if (!is.null(analyte) && analyte != "") {
      filtered_data <- extracted_data() %>%
        filter(Analyte == analyte)

      # Exclude outliers if the checkbox is checked
      if (exclude_outliers) {
        filtered_data <- filtered_data %>%
          filter(abs(Z_score_Method) < 3)  # Exclude Z-scores greater than 3 or less than -3
      }

      # Create a new column for color based on Z-score
      filtered_data <- filtered_data %>%
        mutate(Z_score_Label = ifelse(abs(Z_score_Method) > 3, "Action Required", "OK"))

      # Create the Z-Score plot with conditional coloring
      p <- ggplot(filtered_data, aes(x = Sample, y = Z_score_Method)) +
        geom_point(aes(color = Z_score_Label), shape = 4, size = 3) +  # Color points based on Z-score
        geom_hline(yintercept = c(-3, 3), color = "red", linetype = "longdash") +
        geom_hline(yintercept = c(-2, 2), color = "orange", linetype = "longdash") +
        geom_hline(yintercept = c(-1, 1), color = "green", linetype = "longdash") +
        labs(title = analyte,
             x = "Sample", y = "Z-Score (Method)") +
        scale_color_manual(values = c("Action Required" = "firebrick1", "OK" = "black"),
                           name = "Z-Score",
                           labels = c("Action Required" = "Z-Score >3", "OK" = "Z-Score <=3")) +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1))

      ggplotly(p)
    }
  })

  # Add in frequency plot
  output$additional_plot <- renderPlotly({
    analyte <- input$analyte_plot

    if (!is.null(analyte) && analyte != "") {
      filtered_data <- extracted_data() %>%
        filter(Analyte == analyte)

      if (nrow(filtered_data) == 0) {
        return(NULL)
      }

      p <- ggplot(filtered_data, aes(x = as.numeric(Your_Result))) +
        geom_histogram(bins = 30, fill = "skyblue", color = "black") +
        labs(x = "Your Method Results",
             y = "Frequency") +
        theme_minimal()

      ggplotly(p)
    }
  })


  # Download handlers- Excel and PDF
  # Download Excel report
  output$download_report <- downloadHandler(
    filename = function() { paste("EQA_Report_", Sys.Date(), ".xlsx", sep = "") },
    content = function(file) {
      req(input$analyte_excel)

      # Create a new workbook
      wb <- openxlsx::createWorkbook()

      # Sanitize sheet names to avoid colons or other invalid characters
      sanitize_name <- function(name) {
        # Replace colons with an underscore
        gsub("[:]", "_", name)
      }

      # Modify the workbook creation loop to sanitize analyte names
      for (analyte in input$analyte_excel) {
        # Filter the data for the current analyte
        selected_data <- filter(extracted_data(), Analyte == analyte)

        # Sanitize the analyte name for the worksheet
        sanitized_analyte <- sanitize_name(analyte)

        # Add a new worksheet with the sanitized analyte name
        openxlsx::addWorksheet(wb, sanitized_analyte)

        # Write the data to the new sheet
        openxlsx::writeData(wb, sheet = sanitized_analyte, selected_data)
      }

      # Save the workbook to the specified file
      openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
    }
  )

  # Download PDF report
  output$download_summary_report <- downloadHandler(
    filename = function() { paste("Summary_Report_", Sys.Date(), ".pdf", sep = "") },
    content = function(file) {
      # Save the file as PDF
      pdf(file, width = 8.3, height = 11.7)  # Landscape orientation (A4 size in inches)

      # Loop through selected analytes
      selected_analytes <- input$analyte_pdf

      for (analyte in selected_analytes) {
        # Filter the data based on the selected analyte
        filtered_data <- extracted_data() %>%
          filter(Analyte == analyte)

        # Ensure correct data types for the table
        filtered_data <- filtered_data %>%
          mutate(
            Bias_Method = as.numeric(Bias_Method),
            Your_Result = as.numeric(Your_Result),
            Z_score_Method = as.numeric(Z_score_Method)
          ) %>%
          filter(!is.na(Bias_Method), !is.na(Your_Result), !is.na(Z_score_Method))

        # Create a page with the plots
        grid.newpage()
        # Add header with the analyte name (move title lower to avoid overlap with the plots)
        header_text <- paste(analyte, ": EQA Return Review")
        grid.text(header_text, x = 0.5, y = 0.975, gp = gpar(fontsize = 14, fontface = "bold"))  # Center-aligned header

        # Draw tick box on the right side (adjusted position to align neatly with the header)
        grid.rect(x = 0.95, y = 0.96, width = 0.03, height = 0.03, just = "right",
                  gp = gpar(fill = "white", col = "black"))

        # Add text label next to the tick box on the right side (adjusted to align with tick box)
        grid.text("Checked?", x = 0.9, y = 0.96, just = "right",
                  gp = gpar(fontsize = 10, fontface = "bold"))

        # Define the layout for the plots
        pushViewport(viewport(layout = grid.layout(2, 1)))  # 2 rows, 2 columns

        # 1. Regression Plot
        # Fit the linear model
        lm_model <- lm(Mean_Method ~ Your_Result, data = filtered_data)

        # Extract R-squared value
        r_squared <- summary(lm_model)$r.squared

        # Updated Regression Plot with R-squared annotation
        regression_plot <- ggplot(filtered_data, aes(x = Your_Result, y = Mean_Method)) +
          geom_point(color = "blue", size = 2, shape = 16) +  # Smaller points
          geom_smooth(method = "lm", color = "red") +
          labs(x = paste("Your Result (", analyte, "), umol/L", sep = ""), y = "Mean Result (Method Group), umol/L") +

          theme_minimal(base_size = 9) +  # Smaller font size
          theme(axis.title = element_text(size = 9, face = "bold"),
                axis.text.x = element_text(angle = 45, hjust = 1, face = "bold"),
                axis.title.y = element_text(face = "bold"),
                legend.position = "none",
                plot.margin = margin(60, 30, 30, 30),
                panel.grid.major = element_line(color = "grey", size = 0.5),  # Major gridlines
                panel.grid.minor = element_line(color = "lightgrey", size = 0.25)) +  # Minor gridlines
          annotate("text",
                   x = min(filtered_data$Your_Result),  # Left-most x value
                   y = max(filtered_data$Mean_Method), # Top-most y value
                   label = paste("R = ", format(r_squared, digits = 3)),
                   hjust = 0, vjust = 1,  # Align top-left
                   size = 3, color = "black", fontface = "bold")  # Styling

        # Draw Regression Plot in the first cell of the grid (row 1, col 1)
        print(regression_plot, vp = viewport(layout.pos.row = 1, layout.pos.col = 1))

        # 2. Bias Plot
        bias_plot <- ggplot(filtered_data, aes(x = Sample, y = Bias_Method)) +
          geom_point(aes(color = ifelse(Bias_Method > 15, "red", ifelse(Bias_Method < -15, "red", "black"))),
                     shape = 4, size = 3) +  # Smaller points
          geom_line(color = "green") +  # Line in green
          geom_hline(yintercept = c(-10, 10), color = "red", linetype = "longdash") +
          geom_hline(yintercept = 0, color = "black", linetype = "dashed") +  # Horizontal line at y=0
          labs(x = "", y = "Bias (%, Method Group)") +
          scale_color_manual(values = c("red" = "red", "black" = "black")) +
          theme_minimal(base_size = 9) +
          theme(axis.title = element_text(size = 9, face = "bold"),
                axis.text.x = element_text(angle = 45, hjust = 1, face = "bold"),
                axis.title.y = element_text(face = "bold"),
                legend.position = "none",
                plot.margin = margin(30, 30, 30, 30),
                panel.grid.major = element_line(color = "grey", size = 0.5),  # Major gridlines
                panel.grid.minor = element_line(color = "lightgrey", size = 0.25))  # Minor gridlines

        # Draw Bias Plot in the second cell of the grid (row 1, col 2)
        print(bias_plot, vp = viewport(layout.pos.row = 2, layout.pos.col = 1))

        # Page 2 Layout: Z-Score and Frequency Plots
        grid.newpage()

        # Define the layout of the plots
        pushViewport(viewport(layout = grid.layout(2, 1)))

        # 3. Z-Score Plot
        z_score_plot <- ggplot(filtered_data, aes(x = Sample, y = Z_score_Method)) +
          geom_point(aes(color = ifelse(Z_score_Method < 3.5, "black", "red")), shape = 4, size = 3) +  # Color points based on Z-score
          geom_hline(yintercept = c(-3, 3), color = "red", linetype = "longdash") +
          geom_hline(yintercept = c(-2, 2), color = "orange", linetype = "longdash") +
          geom_hline(yintercept = c(-1, 1), color = "green", linetype = "longdash") +
          labs(x = "", y = "Z-Score (Method Group)") +
          scale_color_manual(values = c("black" = "black", "red" = "red")) +  # Ensure that black and red colors are used
          theme_minimal(base_size = 9) +  # Smaller font size
          theme(axis.title = element_text(size = 9, face = "bold"),
                axis.text.x = element_text(angle = 45, hjust = 1, face = "bold"),
                axis.title.y = element_text(face = "bold"),
                legend.position = "none",
                plot.margin = margin(60, 30, 30, 30),
                panel.grid.major = element_line(color = "grey", size = 0.5),  # Major gridlines
                panel.grid.minor = element_line(color = "lightgrey", size = 0.25))  # Minor gridlines

        # Draw Z-Score Plot in the third cell of the grid (row 2, col 1)
        print(z_score_plot, vp = viewport(layout.pos.row = 1, layout.pos.col = 1))

        # 4. Frequency Plot
        additional_plot <- ggplot(filtered_data, aes(x = as.numeric(Mean_Method))) +
          geom_histogram(bins = 15, fill = "skyblue", color = "black") +
          labs(x = paste(analyte, ", umol/L", sep = ""), y = "Frequency") +
          theme_minimal(base_size = 9) +
          theme(axis.title = element_text(size = 9, face = "bold"),
                axis.text.x = element_text(hjust = 1, face = "bold"),
                axis.title.y = element_text(face = "bold"),
                legend.position = "none",
                plot.margin = margin(30, 30, 30, 30),
                panel.grid.major = element_line(color = "grey", size = 0.5),  # Major gridlines
                panel.grid.minor = element_line(color = "lightgrey", size = 0.25))  # Minor gridlines

        # Draw Frequency Plot in the fourth cell of the grid (row 2, col 2)
        print(additional_plot, vp = viewport(layout.pos.row = 2, layout.pos.col = 1))

        popViewport()  # Close main grid layout viewport
      }

      # Page 3 Layout: EQA Review
      # Path to the logo image
      logo_path <- "synnovis_logo.jpg"
      logo_img <- readJPEG(logo_path)

      # Define the header text
      header_text <- "EQA Return Review"  # Custom header text

      # Create the page with header and logo
      grid.newpage()

      # Add header text in the top-left corner
      grid.text(header_text,
                gp = gpar(fontsize = 16, fontface = "bold"),
                x = 0.05, y = 0.95, just = "left", vjust = 1)

      # Add logo in the top-right corner with adjusted size and position
      pushViewport(viewport(x = 0.95, y = 0.95, width = 0.1, height = 0.1, just = c("right", "top")))
      grid.raster(logo_img, width = 1.1811, height = 0.3149, x = 0.05, y = 0.98)
      popViewport()

      # Create a structured layout with adjusted row heights for better spacing
      pushViewport(viewport(layout = grid.layout(15, 1,
                                                 heights = unit(c(1, 0.5, 0.5, 0.5, 1, 0.5, 0.5, 0.5, 1, 2, 0.5, 0.5, 0.5, 1, 1), "null"))))

      # Section 1: EQA Return Details (with spacing adjustments)
      grid.text("Scheme Provider: _________________________________   Scheme: _________________________________",
                vp = viewport(layout.pos.row = 2, layout.pos.col = 1, just = "left"),
                gp = gpar(fontsize = 11, fontface = "bold"))

      grid.text("Distribution: ________________________________     Method Set: ___________________________________",
                vp = viewport(layout.pos.row = 3, layout.pos.col = 1, just = "left"),
                gp = gpar(fontsize = 11, fontface = "bold"))

      # Section 2: EQA Review - Acknowledgment
      grid.text("Reviewed by: ___________________________________    Date: _____________________________________",
                vp = viewport(layout.pos.row = 4, layout.pos.col = 1, just = "left"),
                gp = gpar(fontsize = 11, fontface = "bold"))


      # Comments Section (adjustments for a more professional layout)
      grid.text("Comments:                                                                                                                                             ",
                vp = viewport(layout.pos.row = 5, layout.pos.col = 1, just = "left"),
                gp = gpar(fontsize = 11, fontface = "bold"))

      grid.rect(x = 0.5, y = 0.15, width = 0.9, height = 1.15, just = "center",
                gp = gpar(col = "black", fill = NA))  # Centered and larger box for comments

      # Footer Section (professional and aligned to the left)
      grid.text("Please ensure all fields are completed before submission. \nIf further investigation is required, please complete the IMD EQA Troubleshooting Checklist (SLMM-LF-96).",
                vp = viewport(layout.pos.row = 15, layout.pos.col = 1, width = 0.8,
                              just = "left"),
                gp = gpar(fontsize = 9, fontface = "italic"))

      # End the layout viewport
      popViewport()


      dev.off()  # Close the PDF device
    },
    contentType = "application/pdf"
  )
}

# Run the application
shinyApp(ui = ui, server = server)
