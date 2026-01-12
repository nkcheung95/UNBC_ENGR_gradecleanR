# Grade Summary Data Cleaning Script v1.4
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
    
    if (nrow(student_data) == 0) return(NULL)
    
    student_numbers <- student_data[[1]]
    scores <- student_data[, -1] 
    n_assignments <- ncol(scores)
    assignment_cols <- paste0("Assignment_", 1:n_assignments)
    
    wide_df <- data.frame(student_number = student_numbers, scores, stringsAsFactors = FALSE)
    colnames(wide_df) <- c("student_number", assignment_cols)
    
    long_df <- wide_df %>%
      pivot_longer(cols = starts_with("Assignment_"), names_to = "assignment", values_to = "score") %>%
      mutate(assignment_num = as.numeric(str_extract(assignment, "\\d+"))) %>%
      mutate(
        course_name = course_name,
        attribute = attribute[assignment_num],
        indicator = indicator[assignment_num],
        level_assessed = level_assessed[assignment_num],
        student_id = paste(course_name, student_number, sep = "_")
      ) %>%
      select(student_id, course_name, student_number, attribute, indicator, level_assessed, score) %>%
      filter(!is.na(score)) 
    
    return(long_df)
    
  }, error = function(e) {
    return(list(is_error = TRUE, file = basename(file_path), message = conditionMessage(e)))
  })
}

# Process All Files
# -----------------
all_results <- lapply(file_paths, process_grade_file)

failed_files <- list()
successful_data <- list()

for (result in all_results) {
  if (is.null(result)) next
  # Fix for the "Unknown column error" warning:
  if (is.list(result) && "is_error" %in% names(result)) {
    failed_files[[length(failed_files) + 1]] <- result
  } else {
    successful_data[[length(successful_data) + 1]] <- result
  }
}

if (length(successful_data) == 0) stop("No data successfully processed.")

combined_data <- bind_rows(successful_data)
combined_data$course_name <- factor(combined_data$course_name)
combined_data$attribute <- factor(combined_data$attribute, levels = unique(combined_data$attribute)[order(as.numeric(unique(combined_data$attribute)))])
combined_data$indicator <- factor(combined_data$indicator, levels = unique(combined_data$indicator)[order(as.numeric(unique(combined_data$indicator)))])
combined_data$level_assessed <- factor(combined_data$level_assessed, levels = c("I", "D", "A"))
combined_data$score_label <- factor(combined_data$score, levels = c("4", "3", "2", "1"), labels = c("Fail", "Minimal", "Adequate", "Exceeds"))

# Plotting Functions with Dynamic Wrapping
# ----------------------------------------

# Logic: Narrower plots (fewer items) need more aggressive wrapping
get_dynamic_wrap <- function(n_items) {
  if (n_items <= 2) return(25)
  if (n_items <= 5) return(40)
  return(60)
}

create_stacked_bar <- function(data, title, type = "indicator") {
  col_sym <- sym(type)
  n_items <- length(unique(data[[type]]))
  wrap_width <- get_dynamic_wrap(n_items)
  
  plot_data <- data %>%
    group_by(!!col_sym, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(!!col_sym) %>%
    mutate(percentage = count / sum(count) * 100)
  
  summary_stats <- data %>%
    group_by(!!col_sym) %>%
    summarise(
      m = median(as.numeric(score), na.rm = TRUE),
      se = sd(as.numeric(score), na.rm = TRUE) / sqrt(n()),
      label = sprintf("M: %.1f\nSE: %.2f", m, se),
      .groups = 'drop'
    )

  ggplot(plot_data, aes(x = !!col_sym, y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><br><span style='font-size:7pt'>(", sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5), color = "white", fill = NA, label.color = NA, size = 3) +
    geom_text(data = summary_stats, aes(x = !!col_sym, y = 105, label = label), inherit.aes = FALSE, size = 3.5, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F", "Minimal" = "#FFA726", "Adequate" = "#66BB6A", "Exceeds" = "#2E7D32")) +
    labs(title = str_wrap(title, width = wrap_width), x = str_to_title(gsub("_", " ", type)), y = "Percentage (%)", fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 14) +
    theme(plot.title = element_text(hjust = 0.5, face = "bold", lineheight = 1.1),
          axis.text.x = element_text(angle = 45, hjust = 1))
}

# Directory Setup
# ---------------
main_output_dir <- file.path(dirname(file_paths[1]), "grade_outputs")
dir.create(main_output_dir, showWarnings = FALSE, recursive = TRUE)

# Generate Amalgamated Plots
# --------------------------
cat("Generating Amalgamated Plots...\n")
ggsave(file.path(main_output_dir, "amalgamated_indicator.png"), create_stacked_bar(combined_data, "All Courses - Score Distribution by Indicator", "indicator"), width = 10, height = 7)
ggsave(file.path(main_output_dir, "amalgamated_attribute.png"), create_stacked_bar(combined_data, "All Courses - Score Distribution by Attribute", "attribute"), width = 10, height = 7)
ggsave(file.path(main_output_dir, "amalgamated_level.png"), create_stacked_bar(combined_data, "All Courses - Score Distribution by Level Assessed", "level_assessed"), width = 8, height = 7)

# Generate Course-Specific Plots
# ------------------------------
courses <- unique(combined_data$course_name)
for (course in courses) {
  cat(paste("Processing Course:", course, "\n"))
  course_clean <- gsub("[ .:]", "_", course)
  course_dir <- file.path(main_output_dir, course_clean)
  dir.create(course_dir, showWarnings = FALSE)
  
  course_sub <- combined_data %>% filter(course_name == course)
  
  # Indicator
  ggsave(file.path(course_dir, paste0(course_clean, "_indicator.png")), 
         create_stacked_bar(course_sub, paste(course, "- Score Distribution by Indicator"), "indicator"), width = 9, height = 7)
  # Attribute
  ggsave(file.path(course_dir, paste0(course_clean, "_attribute.png")), 
         create_stacked_bar(course_sub, paste(course, "- Score Distribution by Attribute"), "attribute"), width = 9, height = 7)
  # Level Assessed
  ggsave(file.path(course_dir, paste0(course_clean, "_level.png")), 
         create_stacked_bar(course_sub, paste(course, "- Score Distribution by Level Assessed"), "level_assessed"), width = 7, height = 7)
}

cat("\nDone! All outputs saved in 'grade_outputs'.\n")
