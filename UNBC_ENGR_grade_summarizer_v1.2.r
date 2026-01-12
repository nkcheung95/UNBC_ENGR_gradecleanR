# Grade Summary Data Cleaning Script v1.3
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
    
    raw_data <- read_excel(file_path, sheet = 1, col_names = FALSE)
    course_name <- as.character(raw_data[1, 1])
    
    attribute_raw <- as.character(raw_data[6, -1]) 
    indicator_raw <- as.character(raw_data[7, -1])
    level_assessed <- as.character(raw_data[8, -1])
    
    attribute <- str_extract(attribute_raw, "\\d+")
    indicator <- str_extract(indicator_raw, "\\d+\\.?\\d*")
    student_data <- raw_data[11:nrow(raw_data), ]
    student_data <- student_data[!is.na(student_data[[1]]), ]
    
    if (nrow(student_data) == 0) {
      cat("  Warning: No student data found\n")
      return(NULL)
    }
    
    student_numbers <- student_data[[1]]
    scores <- student_data[, -1] 
    
    n_assignments <- ncol(scores)
    assignment_cols <- paste0("Assignment_", 1:n_assignments)
    
    wide_df <- data.frame(
      student_number = student_numbers,
      scores,
      stringsAsFactors = FALSE
    )
    colnames(wide_df) <- c("student_number", assignment_cols)
    
    long_df <- wide_df %>%
      pivot_longer(
        cols = starts_with("Assignment_"),
        names_to = "assignment",
        values_to = "score"
      ) %>%
      mutate(
        assignment_num = as.numeric(str_extract(assignment, "\\d+"))
      )
    
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
      filter(!is.na(score)) 
    
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

if (length(successful_data) == 0) {
  stop("No data to process. Exiting.")
}

combined_data <- bind_rows(successful_data)
combined_data$course_name <- factor(combined_data$course_name)
combined_data$attribute <- factor(combined_data$attribute, 
                                  levels = unique(combined_data$attribute)[order(as.numeric(unique(combined_data$attribute)))])
combined_data$indicator <- factor(combined_data$indicator, 
                                  levels = unique(combined_data$indicator)[order(as.numeric(unique(combined_data$indicator)))])
combined_data$level_assessed <- factor(combined_data$level_assessed, 
                                       levels = c("I", "D", "A"))

# Clean filenames helper
clean_filename <- function(name) {
  name <- gsub("[ .:]", "_", name)
  name <- gsub("_{2,}", "_", name)
  name <- gsub("^_|_$", "", name)
  return(name)
}

# Visualization Functions (Updated with Dynamic Wrapping)
# =======================================================

# Helper to determine wrap width based on plot width
get_wrap_width <- function(n_elements, is_isolated = FALSE) {
  if (is_isolated) return(25)
  return(max(30, n_elements * 6)) # Narrower plots get tighter wrapping
}

create_stacked_bar <- function(data, title) {
  n_ind <- length(unique(data$indicator))
  wrap_val <- get_wrap_width(n_ind)
  
  plot_data <- data %>%
    group_by(indicator, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(indicator) %>%
    mutate(percentage = count / sum(count) * 100) %>%
    ungroup() %>%
    mutate(indicator_num = as.numeric(as.character(indicator)))
  
  summary_stats <- data %>%
    group_by(indicator) %>%
    summarise(
      median_score = median(as.numeric(score), na.rm = TRUE),
      se_score = sd(as.numeric(score), na.rm = TRUE) / sqrt(n()),
      label = sprintf("M: %.1f\nSE: %.2f", median_score, se_score),
      indicator_num = as.numeric(as.character(indicator[1])),
      .groups = 'drop'
    )
  
  ggplot(plot_data, aes(x = reorder(indicator, indicator_num), y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(data = summary_stats, aes(x = reorder(indicator, indicator_num), y = 105, label = label),
              inherit.aes = FALSE, size = 3, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F", "Minimal" = "#FFA726", "Adequate" = "#66BB6A", "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(title, width = wrap_val), x = "Indicator", y = "Percentage (%)", fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5, face = "bold", lineheight = 1.1)) #
}

create_stacked_bar_attribute <- function(data, title) {
  n_attr <- length(unique(data$attribute))
  wrap_val <- get_wrap_width(n_attr)
  
  plot_data <- data %>%
    group_by(attribute, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(attribute) %>%
    mutate(percentage = count / sum(count) * 100) %>%
    ungroup() %>%
    mutate(attribute_num = as.numeric(as.character(attribute)))
  
  summary_stats <- data %>%
    group_by(attribute) %>%
    summarise(
      median_score = median(as.numeric(score), na.rm = TRUE),
      se_score = sd(as.numeric(score), na.rm = TRUE) / sqrt(n()),
      label = sprintf("M: %.1f\nSE: %.2f", median_score, se_score),
      attribute_num = as.numeric(as.character(attribute[1])),
      .groups = 'drop'
    )
  
  ggplot(plot_data, aes(x = reorder(attribute, attribute_num), y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(data = summary_stats, aes(x = reorder(attribute, attribute_num), y = 105, label = label),
              inherit.aes = FALSE, size = 3, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F", "Minimal" = "#FFA726", "Adequate" = "#66BB6A", "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(title, width = wrap_val), x = "Attribute", y = "Percentage (%)", fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5, face = "bold", lineheight = 1.1)) #
}

create_stacked_bar_level <- function(data, title) {
  n_lvls <- length(unique(data$level_assessed))
  wrap_val <- get_wrap_width(n_lvls)
  
  plot_data <- data %>%
    group_by(level_assessed, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(level_assessed) %>%
    mutate(percentage = count / sum(count) * 100) %>%
    ungroup()
  
  summary_stats <- data %>%
    group_by(level_assessed) %>%
    summarise(
      median_score = median(as.numeric(score), na.rm = TRUE),
      se_score = sd(as.numeric(score), na.rm = TRUE) / sqrt(n()),
      label = sprintf("M: %.1f\nSE: %.2f", median_score, se_score),
      .groups = 'drop'
    )
  
  ggplot(plot_data, aes(x = level_assessed, y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(data = summary_stats, aes(x = level_assessed, y = 105, label = label),
              inherit.aes = FALSE, size = 3, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F", "Minimal" = "#FFA726", "Adequate" = "#66BB6A", "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(title, width = wrap_val), x = "Level Assessed", y = "Percentage (%)", fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5, face = "bold", lineheight = 1.1)) #
}

# Single Variable Plots (Isolated)
# --------------------------------
create_single_variable_plot <- function(data, val, type = "Indicator") {
  # Universal function for isolated single-bar plots
  filter_col <- if(type == "Indicator") "indicator" else if(type == "Attribute") "attribute" else "level_assessed"
  plot_data_sub <- data[data[[filter_col]] == val, ]
  
  stats <- plot_data_sub %>%
    group_by(score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    mutate(percentage = count / sum(count) * 100)
  
  median_val <- median(as.numeric(plot_data_sub$score), na.rm = TRUE)
  se_val <- sd(as.numeric(plot_data_sub$score), na.rm = TRUE) / sqrt(nrow(plot_data_sub))
  
  ggplot(stats, aes(x = "", y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE), width = 0.7) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><br><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5), color = "white", fill = NA, label.color = NA, size = 2.8) +
    annotate("text", x = 1, y = 115, label = sprintf("M: %.1f\nSE: %.2f", median_val, se_val), size = 3.5) +
    scale_fill_manual(values = c("Fail" = "#D32F2F", "Minimal" = "#FFA726", "Adequate" = "#66BB6A", "Exceeds" = "#2E7D32")) +
    labs(title = str_wrap(paste(type, val), width = 20), x = "", y = "Percentage (%)", fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(plot.title = element_text(hjust = 0.5, face = "bold", lineheight = 1.1)) #
}

# Output Generation
# =================
output_folder <- dirname(file_paths[1])
main_output_dir <- file.path(output_folder, "grade_outputs")
dir.create(main_output_dir, showWarnings = FALSE, recursive = TRUE)

combined_data$score_label <- factor(combined_data$score, 
                                    levels = c("4", "3", "2", "1"),
                                    labels = c("Fail", "Minimal", "Adequate", "Exceeds"))

# Amalgamated Runs
# ----------------
cat("\nSaving amalgamated figures...\n")
ggsave(file.path(main_output_dir, "amalgamated_indicator.png"), create_stacked_bar(combined_data, "All Courses - Score Distribution by Indicator"), width = max(8, length(unique(combined_data$indicator))*0.8), height = 6)
ggsave(file.path(main_output_dir, "amalgamated_attribute.png"), create_stacked_bar_attribute(combined_data, "All Courses - Score Distribution by Attribute"), width = max(8, length(unique(combined_data$attribute))*0.8), height = 6)
ggsave(file.path(main_output_dir, "amalgamated_level.png"), create_stacked_bar_level(combined_data, "All Courses - Score Distribution by Level Assessed"), width = 8, height = 6)

# Course Specific Runs
# --------------------
courses <- unique(combined_data$course_name)
for (course in courses) {
  cat(paste("Processing course:", course, "\n"))
  clean_name <- clean_filename(as.character(course))
  course_folder <- file.path(main_output_dir, clean_name)
  dir.create(course_folder, showWarnings = FALSE)
  
  course_data <- combined_data %>% filter(course_name == course)
  
  ggsave(file.path(course_folder, paste0(clean_name, "_indicator.png")), 
         create_stacked_bar(course_data, paste(course, "- Score Distribution by Indicator")), 
         width = max(8, length(unique(course_data$indicator))*0.8), height = 6)
  
  ggsave(file.path(course_folder, paste0(clean_name, "_attribute.png")), 
         create_stacked_bar_attribute(course_data, paste(course, "- Score Distribution by Attribute")), 
         width = max(8, length(unique(course_data$attribute))*0.8), height = 6)
}

cat("\n=== All outputs saved with dynamic title wrapping ===\n")