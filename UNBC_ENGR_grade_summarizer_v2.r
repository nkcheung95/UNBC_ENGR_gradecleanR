# Grade Summary Data Cleaning Script (Refactored)
# ===============================================

# 1. Setup and Helper Functions
# ---------------------------------
required_packages <- c("ggtext", "ggplot2", "readxl", "dplyr", "tidyr", "tcltk", "stringr", "openxlsx", "rlang")

for (pkg in required_packages) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  }
}

# --- Logging Setup ---
# We define the log file path later after selecting files, but define the function now
log_event <- function(msg, context = "General", type = "INFO") {
  entry <- sprintf("[%s] [%s] %s: %s", Sys.time(), type, context, msg)
  # Write to console
  cat(paste0(entry, "\n"))
  # Write to file (if log_file exists)
  if (exists("log_file")) write(entry, file = log_file, append = TRUE)
}

# --- Generic Clean Filename ---
clean_filename <- function(name) {
  name %>% 
    gsub("[ .:]", "_", .) %>% 
    gsub("_{2,}", "_", .) %>% 
    gsub("^_|_$", "", .)
}

# --- Generic Summary Statistics Function ---
# Replaces the 3 separate summary blocks
calc_stats <- function(df, group_col_name) {
  df %>%
    group_by(across(all_of(group_col_name))) %>%
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
    arrange(across(all_of(group_col_name)))
}

# --- Generic Plotting Function ---
# Replaces create_stacked_bar, create_stacked_bar_attribute, create_stacked_bar_level
plot_distribution <- function(data, group_col_name, title) {
  
  # Ensure the group column is handled as a symbol for plotting
  grp_sym <- sym(group_col_name)
  
  # Prepare Plot Data
  plot_data <- data %>%
    group_by(!!grp_sym, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(!!grp_sym) %>%
    mutate(percentage = count / sum(count) * 100) %>%
    ungroup() 
  
  # Prepare Stats Labels
  summary_stats <- data %>%
    group_by(!!grp_sym) %>%
    summarise(
      mean_score = mean(as.numeric(score), na.rm = TRUE),
      sd_score = sd(as.numeric(score), na.rm = TRUE),
      meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
      label = sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", meet_pct, mean_score, sd_score),
      .groups = 'drop'
    )
  
  # Create Plot
  # Note: sorting logic is handled by factor levels in the main data
  p <- ggplot(plot_data, aes(x = !!grp_sym, y = percentage, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", 
                                     sprintf("%.0f", percentage), "%)</span>")), 
                  position = position_stack(vjust = 0.5),
                  color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(data = summary_stats, aes(x = !!grp_sym, y = 105, label = label),
              inherit.aes = FALSE, size = 2, vjust = 0) +
    scale_fill_manual(values = c("Fail" = "#D32F2F", "Minimal" = "#FFA726",
                                 "Adequate" = "#66BB6A", "Exceeds" = "#2E7D32")) +
    guides(fill = guide_legend(reverse = TRUE)) +
    labs(title = str_wrap(title, width = 40),
         x = str_to_title(gsub("_", " ", group_col_name)),
         y = "Percentage of Students (%)", fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 15) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5, face = "bold"))
  
  return(p)
}

# --- Generic Single Variable Plot (For isolated plots) ---
plot_single_iso <- function(data, group_col, value) {
  sub_data <- data %>% filter(.data[[group_col]] == value)
  
  # Create a dummy column for x-axis since we are plotting a single item
  p <- plot_distribution(sub_data %>% mutate(dummy = ""), "dummy", paste(group_col, value)) +
    theme(axis.text.x = element_blank(), axis.ticks.x = element_blank()) +
    labs(x = NULL)
  
  return(p)
}

# 2. File Selection and Initial Processing
# ----------------------------------------
file_paths <- tk_choose.files(caption = "Select Excel files", filters = matrix(c("Excel", ".xlsx;.xls"), 1, 2))
if (length(file_paths) == 0) stop("No files selected.")

# Initialize Log
log_file <- file.path(dirname(file_paths[1]), "processing_log.txt")
write(paste("Run Date:", Sys.time(), "\n---"), file = log_file)
cat(paste("Selected", length(file_paths), "files.\n"))

# --- File Processing Function ---
process_grade_file <- function(file_path) {
  tryCatch({
    log_event(paste("Reading", basename(file_path)), "FileIO")
    raw_data <- read_excel(file_path, sheet = 1, col_names = FALSE)
    
    # Validation
    if(nrow(raw_data) < 11) return(NULL)
    
    # Metadata extraction
    course_name <- as.character(raw_data[1, 1])
    # Extract numerics/strings
    attr <- str_extract(as.character(raw_data[6, -1]), "\\d+")
    ind  <- str_extract(as.character(raw_data[7, -1]), "\\d+\\.?\\d*")
    lvl  <- as.character(raw_data[8, -1])
    
    # Student Data
    s_data <- raw_data[11:nrow(raw_data), ]
    s_data <- s_data[!is.na(s_data[[1]]), ]
    if(nrow(s_data) == 0) return(NULL)
    
    # Reshape
    colnames(s_data) <- c("student_number", paste0("A_", 1:length(attr)))
    long_df <- s_data %>%
      pivot_longer(cols = starts_with("A_"), names_to = "idx", values_to = "score") %>%
      mutate(
        idx_num = as.numeric(gsub("A_", "", idx)),
        course_name = course_name,
        attribute = attr[idx_num],
        indicator = ind[idx_num],
        level_assessed = lvl[idx_num],
        student_id = paste(course_name, student_number, sep = "_")
      ) %>%
      filter(!is.na(score)) %>%
      select(student_id, course_name, student_number, attribute, indicator, level_assessed, score)
    
    return(long_df)
    
  }, error = function(e) {
    log_event(e$message, basename(file_path), "ERROR")
    return(list(error = TRUE, file = basename(file_path), msg = e$message))
  })
}

# 3. Execution & Aggregation
# --------------------------
all_results <- lapply(file_paths, process_grade_file)

successful_data <- list()
failed_files <- list()

for (res in all_results) {
  if (is.list(res) && !is.null(res$error)) failed_files[[length(failed_files) + 1]] <- res
  else if (!is.null(res)) successful_data[[length(successful_data) + 1]] <- res
}

if (length(successful_data) == 0) stop("No valid data processed.")

combined_data <- bind_rows(successful_data)

# Factor Setup (Global)
combined_data <- combined_data %>%
  mutate(
    course_name = factor(course_name),
    attribute = factor(attribute, levels = sort(unique(as.numeric(attribute)))),
    indicator = factor(indicator, levels = sort(unique(as.numeric(indicator)))),
    level_assessed = factor(level_assessed, levels = c("I", "D", "A")),
    score_label = factor(score, levels = c("4", "3", "2", "1"), 
                         labels = c("Fail", "Minimal", "Adequate", "Exceeds"))
  )

# Output Setup
output_folder <- dirname(file_paths[1])
main_output_dir <- file.path(output_folder, "grade_outputs")
dir.create(main_output_dir, showWarnings = FALSE, recursive = TRUE)
write.csv(combined_data, file.path(main_output_dir, "cleaned_grade_data.csv"), row.names = FALSE)

# 4. Amalgamated Outputs (Vectorized)
# -----------------------------------
log_event("Generating Amalgamated Outputs", "Main")
vars_to_process <- c("attribute", "indicator", "level_assessed")
wb_amalgamated <- createWorkbook()

for (var in vars_to_process) {
  # 1. Summary Stats
  stats <- calc_stats(combined_data, var)
  addWorksheet(wb_amalgamated, paste0("By_", str_to_title(var)))
  writeData(wb_amalgamated, paste0("By_", str_to_title(var)), stats)
  
  # 2. Plots
  # Dynamic width calculation
  n_items <- length(unique(combined_data[[var]]))
  p_width <- if(var == "level_assessed") max(6, n_items * 2) else max(8, n_items * 0.8)
  
  p <- plot_distribution(combined_data, var, paste("All Courses - by", var))
  ggsave(file.path(main_output_dir, paste0("amalgamated_figure_", var, ".png")), 
         p, width = p_width, height = 6, dpi = 300)
}
saveWorkbook(wb_amalgamated, file.path(main_output_dir, "amalgamated_summary.xlsx"), overwrite = TRUE)

# 5. Isolated Plots (Nested Loops)
# --------------------------------
log_event("Generating Isolated Plots", "Main")
iso_dir <- file.path(main_output_dir, "isolated_amalgamated_plots")

for (var in vars_to_process) {
  # Create subfolder (e.g., "attribute_plots")
  sub_dir <- file.path(iso_dir, paste0(var, "_plots"))
  dir.create(sub_dir, recursive = TRUE, showWarnings = FALSE)
  
  unique_vals <- unique(combined_data[[var]])
  unique_vals <- unique_vals[!is.na(unique_vals)]
  
  for (val in unique_vals) {
    # Clean value for filename
    safe_val <- gsub("\\.", "_", as.character(val))
    fname <- file.path(sub_dir, paste0(var, "_", safe_val, ".png"))
    
    p <- plot_single_iso(combined_data, var, val)
    ggsave(fname, p, width = 4, height = 6, dpi = 300)
  }
}

# 6. Course-Specific Processing (Loop with Warning Logging)
# ---------------------------------------------------------
courses <- unique(combined_data$course_name)

for (course in courses) {
  # We use withCallingHandlers to catch warnings without breaking the loop
  withCallingHandlers({
    
    log_event(paste("Processing outputs for:", course), "CourseLoop")
    
    # A. Setup
    clean_name <- clean_filename(as.character(course))
    course_dir <- file.path(main_output_dir, clean_name)
    dir.create(course_dir, showWarnings = FALSE)
    
    course_data <- combined_data %>% filter(course_name == course)
    
    # B. Generate Workbook for this course
    wb_course <- createWorkbook()
    
    # Loop through the 3 variables (Attribute, Indicator, Level)
    for (var in vars_to_process) {
      
      # Statistics
      c_stats <- calc_stats(course_data, var)
      sheet_name <- paste0("By_", str_to_title(var))
      addWorksheet(wb_course, sheet_name)
      writeData(wb_course, sheet_name, c_stats)
      
      # Figure
      n_items <- length(unique(course_data[[var]]))
      p_width <- if(var == "level_assessed") max(6, n_items * 2) else max(8, n_items * 0.8)
      
      c_plot <- plot_distribution(course_data, var, paste(course, "- by", var))
      ggsave(file.path(course_dir, paste0(clean_name, "_figure_", var, ".png")),
             c_plot, width = p_width, height = 6, dpi = 300)
    }
    
    saveWorkbook(wb_course, file.path(course_dir, paste0(clean_name, "_summary.xlsx")), overwrite = TRUE)
    
  }, warning = function(w) {
    # --- WARNING HANDLER ---
    # Log the warning to file with context
    log_event(w$message, paste("Warning in Course:", course), "WARNING")
    # Muffle the warning so it doesn't print to console repeatedly/ugly
    invokeRestart("muffleWarning")
    
  }, error = function(e) {
    # --- ERROR HANDLER ---
    log_event(e$message, paste("Error in Course:", course), "ERROR")
  })
}

# Final Report
# ------------
cat("\n=== Processing Complete ===\n")
cat(paste("Detailed logs saved to:", log_file, "\n"))
if (length(failed_files) > 0) {
  cat("! Some input files failed. Check log for details.\n")
}