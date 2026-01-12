# Grade Summary Data Cleaning Script
# =====================================

# Package Installation and Loading
# ---------------------------------
required_packages <- c("ggtext","ggplot2","readxl", "dplyr", "tidyr", "tcltk", "stringr","openxlsx")

for (pkg in required_packages) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  }
}

# File Selection
# --------------
cat("Select Excel files containing grade data...\n")
file_paths <- tk_choose.files(
  caption = "Select Excel files with grade summaries",
  multi = TRUE,
  filters = matrix(c("Excel files", ".xlsx;.xls"), 1, 2)
)

if (length(file_paths) == 0) {
  stop("No files selected. Exiting.")
}

cat(paste("Selected", length(file_paths), "file(s)\n\n"))

# Data Processing Function
# ------------------------
process_grade_file <- function(file_path) {
  
  tryCatch({
    cat(paste("Processing:", basename(file_path), "\n"))
    
    # Read the entire first sheet
    raw_data <- read_excel(file_path, sheet = 1, col_names = FALSE)
    
    # Extract course name from A1
    course_name <- as.character(raw_data[1, 1])
    
    # Extract metadata from rows 6, 7, 8 (columns B onward)
    attribute_raw <- as.character(raw_data[6, -1])  # Remove column A
    indicator_raw <- as.character(raw_data[7, -1])
    level_assessed <- as.character(raw_data[8, -1])
    
    # Extract numbers from attribute and indicator
    # Attribute: only whole numbers (no decimals)
    attribute <- str_extract(attribute_raw, "\\d+")
    # Indicator: include decimals
    indicator <- str_extract(indicator_raw, "\\d+\\.?\\d*")
    # Extract student data starting from row 11
    student_data <- raw_data[11:nrow(raw_data), ]
    
    # Keep only rows with student numbers (non-NA in column A)
    student_data <- student_data[!is.na(student_data[[1]]), ]
    
    # If no student data, return NULL
    if (nrow(student_data) == 0) {
      cat("  Warning: No student data found\n")
      return(NULL)
    }
    
    # Extract student numbers and scores
    student_numbers <- student_data[[1]]
    scores <- student_data[, -1]  # All columns except A
    
    # Create column names for assignments
    n_assignments <- ncol(scores)
    assignment_cols <- paste0("Assignment_", 1:n_assignments)
    
    # Build wide format dataframe
    wide_df <- data.frame(
      student_number = student_numbers,
      scores,
      stringsAsFactors = FALSE
    )
    colnames(wide_df) <- c("student_number", assignment_cols)
    
    # Transform to long format
    long_df <- wide_df %>%
      pivot_longer(
        cols = starts_with("Assignment_"),
        names_to = "assignment",
        values_to = "score"
      ) %>%
      mutate(
        assignment_num = as.numeric(str_extract(assignment, "\\d+"))
      )
    
    # Add metadata for each assignment
    long_df <- long_df %>%
      mutate(
        course_name = course_name,
        attribute = attribute[assignment_num],
        indicator = indicator[assignment_num],
        level_assessed = level_assessed[assignment_num],
        student_id = paste(course_name, student_number, sep = "_")
      ) %>%
      select(student_id, course_name, student_number, attribute, indicator, 
             level_assessed, score) %>%
      filter(!is.na(score))  # Remove rows with no score
    
    cat(paste("  Processed", length(unique(long_df$student_id)), "students\n"))
    
    return(long_df)
    
  }, error = function(e) {
    cat(paste("  ERROR:", conditionMessage(e), "\n"))
    return(list(error = TRUE, file = basename(file_path), message = conditionMessage(e)))
  })
}

# Process All Files
# -----------------
all_data <- lapply(file_paths, process_grade_file)

# Separate successful results from errors
failed_files <- list()
successful_data <- list()

for (i in seq_along(all_data)) {
  result <- all_data[[i]]
  if (is.list(result) && !is.null(result$error) && result$error) {
    failed_files[[length(failed_files) + 1]] <- result
  } else if (!is.null(result)) {
    successful_data[[length(successful_data) + 1]] <- result
  }
}

# Check if we have any successful data
if (length(successful_data) == 0) {
  cat("\n=== ERROR: No files were successfully processed ===\n")
  if (length(failed_files) > 0) {
    cat("\nFailed files:\n")
    for (fail in failed_files) {
      cat(paste("  -", fail$file, ":", fail$message, "\n"))
    }
  }
  stop("No data to process. Exiting.")
}

# Combine all successful data
combined_data <- bind_rows(successful_data)

# Convert course_name to factor
combined_data$course_name <- factor(combined_data$course_name)

# Order attribute and indicator numerically
combined_data$attribute <- factor(combined_data$attribute, 
                                  levels = unique(combined_data$attribute)[order(as.numeric(unique(combined_data$attribute)))])
combined_data$indicator <- factor(combined_data$indicator, 
                                  levels = unique(combined_data$indicator)[order(as.numeric(unique(combined_data$indicator)))])
# Order level_assessed as I, D, A
combined_data$level_assessed <- factor(combined_data$level_assessed, 
                                       levels = c("I", "D", "A"))
# Summary
# -------
cat("\n=== Data Processing Complete ===\n")
cat(paste("Total records:", nrow(combined_data), "\n"))
cat(paste("Unique students:", length(unique(combined_data$student_id)), "\n"))
cat(paste("Courses:", length(unique(combined_data$course_name)), "\n"))
cat("\nCourse breakdown:\n")
print(table(combined_data$course_name))

cat("\nScore distribution:\n")
print(table(combined_data$score))

# Display first few rows
cat("\nFirst 10 rows of cleaned data:\n")
print(head(combined_data, 10))

# Create Summary Tables
# ---------------------
cat("\n=== Summary Statistics ===\n")

# Summary by Attribute
cat("\n--- Students by Score per Attribute ---\n")
summary_attribute <- combined_data %>%
  group_by(attribute) %>%
  summarise(
    n_total = n(),
    n_score_1 = sum(score == 1, na.rm = TRUE),
    n_score_2 = sum(score == 2, na.rm = TRUE),
    n_score_3 = sum(score == 3, na.rm = TRUE),
    n_score_4 = sum(score == 4, na.rm = TRUE),
    meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
    mean_score = mean(as.numeric(score), na.rm = TRUE),
    sd_score = sd(as.numeric(score), na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  arrange(attribute)
print(summary_attribute)

# Summary by Indicator
cat("\n--- Students by Score per Indicator ---\n")
summary_indicator <- combined_data %>%
  group_by(indicator) %>%
  summarise(
    n_total = n(),
    n_score_1 = sum(score == 1, na.rm = TRUE),
    n_score_2 = sum(score == 2, na.rm = TRUE),
    n_score_3 = sum(score == 3, na.rm = TRUE),
    n_score_4 = sum(score == 4, na.rm = TRUE),
    meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
    mean_score = mean(as.numeric(score), na.rm = TRUE),
    sd_score = sd(as.numeric(score), na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  arrange(indicator)
print(summary_indicator)

# Summary by Level Assessed
cat("\n--- Students by Score per Level Assessed ---\n")
summary_level <- combined_data %>%
  group_by(level_assessed) %>%
  summarise(
    n_total = n(),
    n_score_1 = sum(score == 1, na.rm = TRUE),
    n_score_2 = sum(score == 2, na.rm = TRUE),
    n_score_3 = sum(score == 3, na.rm = TRUE),
    n_score_4 = sum(score == 4, na.rm = TRUE),
    meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
    mean_score = mean(as.numeric(score), na.rm = TRUE),
    sd_score = sd(as.numeric(score), na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  arrange(level_assessed)
print(summary_level)

# Function to clean filenames
clean_filename <- function(name) {
  # Replace spaces, periods, and colons with underscores
  name <- gsub("[ .:]", "_", name)
  # Replace multiple underscores with single underscore
  name <- gsub("_{2,}", "_", name)
  # Remove leading/trailing underscores
  name <- gsub("^_|_$", "", name)
  return(name)
}

# Visualization and Output Organization
# =====================================
# Create output directory structure
output_folder <- dirname(file_paths[1])
main_output_dir <- file.path(output_folder, "grade_outputs")
dir.create(main_output_dir, showWarnings = FALSE, recursive = TRUE)

# Save cleaned data
output_file <- file.path(main_output_dir, "cleaned_grade_data.csv")
write.csv(combined_data, output_file, row.names = FALSE)
cat(paste("\n✓ Data saved to:", output_file, "\n"))

# Define score labels for better visualization
score_labels <- c("1" = "Exceeds", "2" = "Adequate", "3" = "Minimal", "4" = "Fail")
combined_data$score_label <- factor(combined_data$score, 
                                    levels = c("4", "3", "2", "1"),
                                    labels = c("Fail", "Minimal", "Adequate", "Exceeds"))

# Function to create stacked bar plot by indicator
create_stacked_bar <- function(data, title) {
  # Calculate percentages
  plot_data <- data %>%
    group_by(indicator, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(indicator) %>%
    mutate(percentage = count / sum(count) * 100) %>%
    ungroup() %>%
    mutate(indicator_num = as.numeric(as.character(indicator)))
  
  # Calculate mean, SD, and meet% for each indicator
  summary_stats <- data %>%
    group_by(indicator) %>%
    summarise(
      mean_score = mean(as.numeric(score), na.rm = TRUE),
      sd_score = sd(as.numeric(score), na.rm = TRUE),
      meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
      label = sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", meet_pct, mean_score, sd_score),
      indicator_num = as.numeric(as.character(indicator[1])),
      .groups = 'drop'
    )
  
  # Create plot
  p <- ggplot(plot_data, aes(x = reorder(indicator, indicator_num), y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(data = summary_stats, aes(x = reorder(indicator, indicator_num), y = 105, label = label),
              inherit.aes = FALSE, size = 2, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F",
                                 "Minimal" = "#FFA726",
                                 "Adequate" = "#66BB6A",
                                 "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(title, width = 40),
         x = "Indicator",
         y = "Percentage of Students (%)",
         fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5, face = "bold"))
  
  return(p)
}

# Function to create stacked bar plot by attribute
create_stacked_bar_attribute <- function(data, title) {
  # Calculate percentages
  plot_data <- data %>%
    group_by(attribute, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(attribute) %>%
    mutate(percentage = count / sum(count) * 100) %>%
    ungroup() %>%
    mutate(attribute_num = as.numeric(as.character(attribute)))
  
  # Calculate mean, SD, and meet% for each attribute
  summary_stats <- data %>%
    group_by(attribute) %>%
    summarise(
      mean_score = mean(as.numeric(score), na.rm = TRUE),
      sd_score = sd(as.numeric(score), na.rm = TRUE),
      meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
      label = sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", meet_pct, mean_score, sd_score),
      attribute_num = as.numeric(as.character(attribute[1])),
      .groups = 'drop'
    )
  
  # Create plot
  p <- ggplot(plot_data, aes(x = reorder(attribute, attribute_num), y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(data = summary_stats, aes(x = reorder(attribute, attribute_num), y = 105, label = label),
              inherit.aes = FALSE, size = 2, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F",
                                 "Minimal" = "#FFA726",
                                 "Adequate" = "#66BB6A",
                                 "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(title, width = 40),
         x = "Attribute",
         y = "Percentage of Students (%)",
         fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5, face = "bold"))
  
  return(p)
}

# Function to create stacked bar plot by level assessed
create_stacked_bar_level <- function(data, title) {
  # Calculate percentages
  plot_data <- data %>%
    group_by(level_assessed, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(level_assessed) %>%
    mutate(percentage = count / sum(count) * 100) %>%
    ungroup()
  
  # Calculate mean, SD, and meet% for each level
  summary_stats <- data %>%
    group_by(level_assessed) %>%
    summarise(
      mean_score = mean(as.numeric(score), na.rm = TRUE),
      sd_score = sd(as.numeric(score), na.rm = TRUE),
      meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
      label = sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", meet_pct, mean_score, sd_score),
      .groups = 'drop'
    )
  
  # Create plot
  p <- ggplot(plot_data, aes(x = level_assessed, y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(data = summary_stats, aes(x = level_assessed, y = 105, label = label),
              inherit.aes = FALSE, size = 2, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F",
                                 "Minimal" = "#FFA726",
                                 "Adequate" = "#66BB6A",
                                 "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(title, width = 40),
         x = "Level Assessed",
         y = "Percentage of Students (%)",
         fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5, face = "bold"))
  
  return(p)
}
# Amalgamated summary statistics
amalgamated_summary_attribute <- combined_data %>%
  group_by(attribute) %>%
  summarise(
    n_total = n(),
    n_score_1 = sum(score == 1, na.rm = TRUE),
    pct_score_1 = round(sum(score == 1, na.rm = TRUE) / n() * 100, 1),
    n_score_2 = sum(score == 2, na.rm = TRUE),
    pct_score_2 = round(sum(score == 2, na.rm = TRUE) / n() * 100, 1),
    n_score_3 = sum(score == 3, na.rm = TRUE),
    pct_score_3 = round(sum(score == 3, na.rm = TRUE) / n() * 100, 1),
    n_score_4 = sum(score == 4, na.rm = TRUE),
    pct_score_4 = round(sum(score == 4, na.rm = TRUE) / n() * 100, 1),
    meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
    mean_score = mean(as.numeric(score), na.rm = TRUE),
    sd_score = sd(as.numeric(score), na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  arrange(attribute)

amalgamated_summary_indicator <- combined_data %>%
  group_by(indicator) %>%
  summarise(
    n_total = n(),
    n_score_1 = sum(score == 1, na.rm = TRUE),
    pct_score_1 = round(sum(score == 1, na.rm = TRUE) / n() * 100, 1),
    n_score_2 = sum(score == 2, na.rm = TRUE),
    pct_score_2 = round(sum(score == 2, na.rm = TRUE) / n() * 100, 1),
    n_score_3 = sum(score == 3, na.rm = TRUE),
    pct_score_3 = round(sum(score == 3, na.rm = TRUE) / n() * 100, 1),
    n_score_4 = sum(score == 4, na.rm = TRUE),
    pct_score_4 = round(sum(score == 4, na.rm = TRUE) / n() * 100, 1),
    meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
    mean_score = mean(as.numeric(score), na.rm = TRUE),
    sd_score = sd(as.numeric(score), na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  arrange(indicator)

# Format indicator column to show one decimal place
amalgamated_summary_indicator$indicator <- sprintf("%.1f", as.numeric(as.character(amalgamated_summary_indicator$indicator)))

amalgamated_summary_level <- combined_data %>%
  group_by(level_assessed) %>%
  summarise(
    n_total = n(),
    n_score_1 = sum(score == 1, na.rm = TRUE),
    pct_score_1 = round(sum(score == 1, na.rm = TRUE) / n() * 100, 1),
    n_score_2 = sum(score == 2, na.rm = TRUE),
    pct_score_2 = round(sum(score == 2, na.rm = TRUE) / n() * 100, 1),
    n_score_3 = sum(score == 3, na.rm = TRUE),
    pct_score_3 = round(sum(score == 3, na.rm = TRUE) / n() * 100, 1),
    n_score_4 = sum(score == 4, na.rm = TRUE),
    pct_score_4 = round(sum(score == 4, na.rm = TRUE) / n() * 100, 1),
    meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
    mean_score = mean(as.numeric(score), na.rm = TRUE),
    sd_score = sd(as.numeric(score), na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  arrange(level_assessed)

# Save amalgamated summaries
amalgamated_summary_workbook <- createWorkbook()
addWorksheet(amalgamated_summary_workbook, "By_Attribute")
addWorksheet(amalgamated_summary_workbook, "By_Indicator")
addWorksheet(amalgamated_summary_workbook, "By_Level")
writeData(amalgamated_summary_workbook, "By_Attribute", amalgamated_summary_attribute)
writeData(amalgamated_summary_workbook, "By_Indicator", amalgamated_summary_indicator)
writeData(amalgamated_summary_workbook, "By_Level", amalgamated_summary_level)
saveWorkbook(amalgamated_summary_workbook, 
             file.path(main_output_dir, "amalgamated_summary.xlsx"), 
             overwrite = TRUE)

# Create and save amalgamated figures
# Calculate dynamic width for indicator plot
n_indicators <- length(unique(combined_data$indicator))
plot_width <- max(8, n_indicators * 0.8)  # minimum 8 inches, 0.8 inches per indicator

amalgamated_plot_indicator <- create_stacked_bar(combined_data, "All Courses - Score Distribution by Indicator")
ggsave(filename = file.path(main_output_dir, "amalgamated_figure_indicator.png"),
       plot = amalgamated_plot_indicator,
       width = plot_width,
       height = 6,
       dpi = 300)

# Calculate dynamic width for attribute plot
n_attributes <- length(unique(combined_data$attribute))
plot_width_attr <- max(8, n_attributes * 0.8)

amalgamated_plot_attribute <- create_stacked_bar_attribute(combined_data, "All Courses - Score Distribution by Attribute")
ggsave(filename = file.path(main_output_dir, "amalgamated_figure_attribute.png"),
       plot = amalgamated_plot_attribute,
       width = plot_width_attr,
       height = 6,
       dpi = 300)

# Calculate dynamic width for level plot
n_levels <- length(unique(combined_data$level_assessed))
plot_width_level <- max(6, n_levels * 2)  # minimum 6 inches, 2 inches per level

amalgamated_plot_level <- create_stacked_bar_level(combined_data, "All Courses - Score Distribution by Level Assessed")
ggsave(filename = file.path(main_output_dir, "amalgamated_figure_level.png"),
       plot = amalgamated_plot_level,
       width = plot_width_level,
       height = 6,
       dpi = 300)
cat("✓ Amalgamated summaries and figures saved\n")

# Create Isolated Plots for Each Variable
# ========================================
cat("\nCreating isolated plots for each variable...\n")

# Create isolated plots directory structure
isolated_plots_dir <- file.path(main_output_dir, "isolated_amalgamated_plots")
indicator_plots_dir <- file.path(isolated_plots_dir, "indicator_plots")
attribute_plots_dir <- file.path(isolated_plots_dir, "attribute_plots")
level_plots_dir <- file.path(isolated_plots_dir, "level_plots")

dir.create(isolated_plots_dir, showWarnings = FALSE, recursive = TRUE)
dir.create(indicator_plots_dir, showWarnings = FALSE, recursive = TRUE)
dir.create(attribute_plots_dir, showWarnings = FALSE, recursive = TRUE)
dir.create(level_plots_dir, showWarnings = FALSE, recursive = TRUE)

# Function to create single variable plot (indicator)
create_single_indicator_plot <- function(data, indicator_val) {
  indicator_data <- data %>% filter(indicator == indicator_val)
  
  plot_data <- indicator_data %>%
    group_by(score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    mutate(percentage = count / sum(count) * 100)
  
  mean_score <- mean(as.numeric(indicator_data$score), na.rm = TRUE)
  sd_score <- sd(as.numeric(indicator_data$score), na.rm = TRUE)
  meet_pct <- round((sum(indicator_data$score %in% c(1, 2, 3), na.rm = TRUE) / nrow(indicator_data)) * 100, 1)
  stats_label <- sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", meet_pct, mean_score, sd_score)
  
  p <- ggplot(plot_data, aes(x = "", y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE), width = 0.7) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(aes(x = "", y = 105, label = stats_label),
              inherit.aes = FALSE, size = 2, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F",
                                 "Minimal" = "#FFA726",
                                 "Adequate" = "#66BB6A",
                                 "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(paste("Indicator", indicator_val), width = 40),
         x = "",
         y = "Percentage of Students (%)",
         fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_blank(),
          axis.ticks.x = element_blank(),
          plot.title = element_text(hjust = 0.5, face = "bold"))
  
  return(p)
}

# Function to create single variable plot (attribute)
create_single_attribute_plot <- function(data, attribute_val) {
  attribute_data <- data %>% filter(attribute == attribute_val)
  
  plot_data <- attribute_data %>%
    group_by(score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    mutate(percentage = count / sum(count) * 100)
  
  mean_score <- mean(as.numeric(attribute_data$score), na.rm = TRUE)
  sd_score <- sd(as.numeric(attribute_data$score), na.rm = TRUE)
  meet_pct <- round((sum(attribute_data$score %in% c(1, 2, 3), na.rm = TRUE) / nrow(attribute_data)) * 100, 1)
  stats_label <- sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", meet_pct, mean_score, sd_score)
  
  p <- ggplot(plot_data, aes(x = "", y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE), width = 0.7) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(aes(x = "", y = 105, label = stats_label),
              inherit.aes = FALSE, size = 2, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F",
                                 "Minimal" = "#FFA726",
                                 "Adequate" = "#66BB6A",
                                 "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(paste("Attribute", attribute_val), width = 40),
         x = "",
         y = "Percentage of Students (%)",
         fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_blank(),
          axis.ticks.x = element_blank(),
          plot.title = element_text(hjust = 0.5, face = "bold"))
  
  return(p)
}

# Function to create single variable plot (level)
create_single_level_plot <- function(data, level_val) {
  level_data <- data %>% filter(level_assessed == level_val)
  
  plot_data <- level_data %>%
    group_by(score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    mutate(percentage = count / sum(count) * 100)
  
  mean_score <- mean(as.numeric(level_data$score), na.rm = TRUE)
  sd_score <- sd(as.numeric(level_data$score), na.rm = TRUE)
  meet_pct <- round((sum(level_data$score %in% c(1, 2, 3), na.rm = TRUE) / nrow(level_data)) * 100, 1)
  stats_label <- sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", meet_pct, mean_score, sd_score)
  
  p <- ggplot(plot_data, aes(x = "", y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE), width = 0.7) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(aes(x = "", y = 105, label = stats_label),
              inherit.aes = FALSE, size = 2, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F",
                                 "Minimal" = "#FFA726",
                                 "Adequate" = "#66BB6A",
                                 "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(paste("Level", level_val), width = 40),
         x = "",
         y = "Percentage of Students (%)",
         fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_blank(),
          axis.ticks.x = element_blank(),
          plot.title = element_text(hjust = 0.5, face = "bold"))
  
  return(p)
}
# Create isolated indicator plots
all_indicators <- unique(combined_data$indicator)
for (ind in all_indicators) {
  ind_plot <- create_single_indicator_plot(combined_data, ind)
  filename <- paste0("indicator_", gsub("\\.", "_", as.character(ind)), ".png")
  ggsave(filename = file.path(indicator_plots_dir, filename),
         plot = ind_plot,
         width = 4,
         height = 6,
         dpi = 300)
  cat(paste("  ✓ Saved:", filename, "\n"))
}

# Create isolated attribute plots
all_attributes <- unique(combined_data$attribute)
for (attr in all_attributes) {
  attr_plot <- create_single_attribute_plot(combined_data, attr)
  filename <- paste0("attribute_", as.character(attr), ".png")
  ggsave(filename = file.path(attribute_plots_dir, filename),
         plot = attr_plot,
         width = 4,
         height = 6,
         dpi = 300)
  cat(paste("  ✓ Saved:", filename, "\n"))
}

# Create isolated level plots
all_levels <- unique(combined_data$level_assessed)
for (lvl in all_levels) {
  if (!is.na(lvl)) {
    lvl_plot <- create_single_level_plot(combined_data, lvl)
    filename <- paste0("level_", as.character(lvl), ".png")
    ggsave(filename = file.path(level_plots_dir, filename),
           plot = lvl_plot,
           width = 4,
           height = 6,
           dpi = 300)
    cat(paste("  ✓ Saved:", filename, "\n"))
  }
}

cat("\n✓ All isolated plots saved\n")
cat(paste("Indicator plots:", length(all_indicators), "files\n"))
cat(paste("Attribute plots:", length(all_attributes), "files\n"))
cat(paste("Level plots:", length(all_levels[!is.na(all_levels)]), "files\n"))

# Course-Specific Summaries and Figures
# -------------------------------------
courses <- unique(combined_data$course_name)

for (course in courses) {
  cat(paste("Processing course:", course, "\n"))
  
  # Clean course name for filenames
  clean_course_name <- clean_filename(as.character(course))
  
  # Create course folder
  course_folder <- file.path(main_output_dir, clean_course_name)
  dir.create(course_folder, showWarnings = FALSE, recursive = TRUE)
  
  # Filter data for this course
  course_data <- combined_data %>% filter(course_name == course)
  
  # Course summary statistics
  course_summary_attribute <- course_data %>%
    group_by(attribute) %>%
    summarise(
      n_total = n(),
      n_score_1 = sum(score == 1, na.rm = TRUE),
      pct_score_1 = round(sum(score == 1, na.rm = TRUE) / n() * 100, 1),
      n_score_2 = sum(score == 2, na.rm = TRUE),
      pct_score_2 = round(sum(score == 2, na.rm = TRUE) / n() * 100, 1),
      n_score_3 = sum(score == 3, na.rm = TRUE),
      pct_score_3 = round(sum(score == 3, na.rm = TRUE) / n() * 100, 1),
      n_score_4 = sum(score == 4, na.rm = TRUE),
      pct_score_4 = round(sum(score == 4, na.rm = TRUE) / n() * 100, 1),
      meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
      mean_score = mean(as.numeric(score), na.rm = TRUE),
      sd_score = sd(as.numeric(score), na.rm = TRUE),
      .groups = 'drop'
    ) %>%
    arrange(attribute)
  
  course_summary_indicator <- course_data %>%
    group_by(indicator) %>%
    summarise(
      n_total = n(),
      n_score_1 = sum(score == 1, na.rm = TRUE),
      pct_score_1 = round(sum(score == 1, na.rm = TRUE) / n() * 100, 1),
      n_score_2 = sum(score == 2, na.rm = TRUE),
      pct_score_2 = round(sum(score == 2, na.rm = TRUE) / n() * 100, 1),
      n_score_3 = sum(score == 3, na.rm = TRUE),
      pct_score_3 = round(sum(score == 3, na.rm = TRUE) / n() * 100, 1),
      n_score_4 = sum(score == 4, na.rm = TRUE),
      pct_score_4 = round(sum(score == 4, na.rm = TRUE) / n() * 100, 1),
      meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
      mean_score = mean(as.numeric(score), na.rm = TRUE),
      sd_score = sd(as.numeric(score), na.rm = TRUE),
      .groups = 'drop'
    ) %>%
    arrange(indicator)
  
  # Format indicator column to show one decimal place
  course_summary_indicator$indicator <- sprintf("%.1f", as.numeric(as.character(course_summary_indicator$indicator)))
  
  course_summary_level <- course_data %>%
    group_by(level_assessed) %>%
    summarise(
      n_total = n(),
      n_score_1 = sum(score == 1, na.rm = TRUE),
      pct_score_1 = round(sum(score == 1, na.rm = TRUE) / n() * 100, 1),
      n_score_2 = sum(score == 2, na.rm = TRUE),
      pct_score_2 = round(sum(score == 2, na.rm = TRUE) / n() * 100, 1),
      n_score_3 = sum(score == 3, na.rm = TRUE),
      pct_score_3 = round(sum(score == 3, na.rm = TRUE) / n() * 100, 1),
      n_score_4 = sum(score == 4, na.rm = TRUE),
      pct_score_4 = round(sum(score == 4, na.rm = TRUE) / n() * 100, 1),
      meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
      mean_score = mean(as.numeric(score), na.rm = TRUE),
      sd_score = sd(as.numeric(score), na.rm = TRUE),
      .groups = 'drop'
    ) %>%
    arrange(level_assessed)
  # Save course summaries
  course_summary_workbook <- createWorkbook()
  addWorksheet(course_summary_workbook, "By_Attribute")
  addWorksheet(course_summary_workbook, "By_Indicator")
  addWorksheet(course_summary_workbook, "By_Level")
  writeData(course_summary_workbook, "By_Attribute", course_summary_attribute)
  writeData(course_summary_workbook, "By_Indicator", course_summary_indicator)
  writeData(course_summary_workbook, "By_Level", course_summary_level)
  saveWorkbook(course_summary_workbook, 
               file.path(course_folder, paste0(clean_course_name, "_summary.xlsx")), 
               overwrite = TRUE)
  
  # Calculate dynamic widths for this course
  n_indicators_course <- length(unique(course_data$indicator))
  n_attributes_course <- length(unique(course_data$attribute))
  n_levels_course <- length(unique(course_data$level_assessed))
  
  # Create and save course figures
  course_plot_indicator <- create_stacked_bar(course_data, 
                                              paste(course, "- Score Distribution by Indicator"))
  ggsave(filename = file.path(course_folder, paste0(clean_course_name, "_figure_indicator.png")),
         plot = course_plot_indicator,
         width = max(8, n_indicators_course * 0.8),
         height = 6,
         dpi = 300)
  
  course_plot_attribute <- create_stacked_bar_attribute(course_data, 
                                                        paste(course, "- Score Distribution by Attribute"))
  ggsave(filename = file.path(course_folder, paste0(clean_course_name, "_figure_attribute.png")),
         plot = course_plot_attribute,
         width = max(8, n_attributes_course * 0.8),
         height = 6,
         dpi = 300)
  
  course_plot_level <- create_stacked_bar_level(course_data, 
                                                paste(course, "- Score Distribution by Level Assessed"))
  ggsave(filename = file.path(course_folder, paste0(clean_course_name, "_figure_level.png")),
         plot = course_plot_level,
         width = max(6, n_levels_course * 2),
         height = 6,
         dpi = 300)
  
  cat(paste("  ✓ Course folder created:", course_folder, "\n"))
}

cat("\n=== All outputs saved ===\n")
cat(paste("Main output directory:", main_output_dir, "\n"))
cat(paste("Total courses processed:", length(courses), "\n"))

# Report failed files at the end
if (length(failed_files) > 0) {
  cat("\n=== WARNING: Some files failed to process ===\n")
  cat(paste("Total failed files:", length(failed_files), "\n\n"))
  for (fail in failed_files) {
    cat(paste("✗ File:", fail$file, "\n"))
    cat(paste("  Error:", fail$message, "\n\n"))
  }
} else {
  cat("\n✓ All files processed successfully!\n")
}
