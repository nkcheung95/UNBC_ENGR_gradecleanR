# Grade Summary Data Cleaning Script (Optimized)
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

# --- Logger Setup ---
log_event <- function(msg, context = "General", type = "INFO") {
  entry <- sprintf("[%s] [%s] {%s}: %s", Sys.time(), type, context, msg)
  if (exists("log_file")) write(entry, file = log_file, append = TRUE)
  # Only print Errors/Info to console, suppress Warnings from console as requested
  if (type != "WARNING") cat(paste0(entry, "\n"))
}

# --- Generic Clean Filename ---
clean_filename <- function(name) {
  name %>% 
    gsub("[^A-Za-z0-9_-]", "_", .) %>%  # Keep only safe characters
    gsub("_{2,}", "_", .) %>% 
    gsub("^_|_$", "", .) %>%
    substr(1, 100)  # Limit length
}
# Line 18-23 (insert after clean_filename function)
# --- Extract Course Code (first 3 digits) ---
extract_course_code <- function(course_name) {
  match <- str_extract(course_name, "^.*?\\d{3}")
  if (is.na(match)) return(clean_filename(course_name))
  return(clean_filename(match))
}
# --- Generic Stats Calculator ---
calc_stats <- function(df, group_col) {
  df %>%
    group_by(across(all_of(group_col))) %>%
    summarise(
      n_total = n(),
      meet_pct = round((sum(score %in% c(1, 2, 3), na.rm = TRUE) / n()) * 100, 1),
      mean_score = mean(as.numeric(score), na.rm = TRUE),
      sd_score = sd(as.numeric(score), na.rm = TRUE),
      n_score_1 = sum(score == 1, na.rm = TRUE),
      n_score_2 = sum(score == 2, na.rm = TRUE),
      n_score_3 = sum(score == 3, na.rm = TRUE),
      n_score_4 = sum(score == 4, na.rm = TRUE),
      .groups = 'drop'
    ) %>%
    arrange(across(all_of(group_col)))
}

# --- Generic Plotter ---
plot_distribution <- function(data, group_col, title, footnote = NULL) {
  grp_sym <- sym(group_col)
  
  # Stats for labels
  stats <- data %>%
    group_by(!!grp_sym) %>%
    summarise(
      label = sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", 
                      sum(score %in% 1:3)/n()*100, mean(as.numeric(score)), sd(as.numeric(score))),
      .groups = 'drop'
    )
  
  # Main Plot
  p <- data %>%
    group_by(!!grp_sym, score_label) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(!!grp_sym) %>%
    mutate(pct = count / sum(count) * 100) %>%
    ggplot(aes(x = !!grp_sym, y = pct, fill = score_label)) +
    geom_bar(stat = "identity", position = position_stack(reverse = TRUE)) +
    geom_richtext(aes(label = paste0("<b>", count, "</b><span style='font-size:7pt'>(", sprintf("%.0f", pct), "%)</span>")), 
                  position = position_stack(vjust = 0.5), color = "white", fill = NA, label.color = NA, size = 2.8) +
    geom_text(data = stats, aes(x = !!grp_sym, y = 105, label = label), inherit.aes = FALSE, size = 2, vjust = 0) +
    scale_fill_manual(values = c("4 - Fail"="#D32F2F", "3 - Minimal"="#FFA726", "2 - Adequate"="#66BB6A", "1 - Exceeds"="#2E7D32"), 
                      breaks = c("1 - Exceeds", "2 - Adequate", "3 - Minimal", "4 - Fail")) +
    labs(title = str_wrap(title, 40), x = str_to_title(gsub("_", " ", group_col)), y = "% Students", fill = "Score") +
    coord_cartesian(ylim = c(0, 130)) +
    theme_classic(base_size = 14) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1), plot.title = element_text(hjust = 0.5, face = "bold"))
  
  # Add footnote if provided
  if (!is.null(footnote)) {
    p <- p + labs(caption = footnote) + 
      theme(plot.caption = element_text(hjust = 0.5, size = 9, color = "gray30"))
  }
  
  return(p)
}

# --- Master Output Generator (The Loop Shortener) ---
# This function handles the Excel + 3 Plots logic for ANY dataset (Combined or Single Course)
generate_output_set <- function(data, folder_path, file_prefix, title_prefix, context_str, footnote = NULL) {
  
  wb <- createWorkbook()
  vars <- c("attribute", "indicator", "level_assessed")
  
  for (var in vars) {
    # 1. Excel Data
    addWorksheet(wb, paste0("By_", str_to_title(var)))
    writeData(wb, paste0("By_", str_to_title(var)), calc_stats(data, var))
    
    # 2. Plotting
    # Dynamic width: Level gets 2 inches per item, others get 0.8 inches
    n_items <- length(unique(data[[var]]))
    p_width <- if(var == "level_assessed") max(6, n_items * 2) else max(8, n_items * 0.8)
    
   p <- plot_distribution(data, var, paste(title_prefix, "-", var), footnote)
    tryCatch({
      ggsave(file.path(folder_path, paste0(file_prefix, "_figure_", var, ".png")), 
             p, width = p_width, height = 6, dpi = 300)
    }, error = function(e) {
      log_event(paste("SKIPPED - ggsave failed:", e$message), 
                paste(context_str, "- figure", var), "ERROR")
    })
  }
  
  saveWorkbook(wb, file.path(folder_path, paste0(file_prefix, "_summary.xlsx")), overwrite = TRUE)
  log_event(paste("Saved outputs to", folder_path), context_str)
}

# 2. File Selection
# -----------------
cat("\n>>> PLEASE SELECT FILES IN THE POP-UP WINDOW <<<\n")
flush.console() # Forces the message to appear immediately

file_paths <- tk_choose.files(caption = "Select Excel Files", filters = matrix(c("Excel", ".xlsx;.xls"), 1, 2))
if (length(file_paths) == 0) stop("No files selected.")

# Extract parent folder name for labeling
parent_folder <- basename(dirname(file_paths[1]))

# Init Log
log_file <- file.path(dirname(file_paths[1]), paste0(parent_folder, "_processing_log.txt"))
write(paste("Run Date:", Sys.time(), "\nParent Folder:", parent_folder, "\n---"), file = log_file)

# 3. Data Ingestion
# -----------------
process_file <- function(fp) {
  tryCatch({
    raw <- read_excel(fp, sheet = 1, col_names = FALSE)
    if(nrow(raw) < 11) return(NULL)
    
    # Extract Metadata & Data
    meta_rows <- list(attr = 6, ind = 7, lvl = 8)
    meta <- lapply(meta_rows, function(r) as.character(raw[r, -1]))
    
    s_data <- raw[11:nrow(raw), ]
    s_data <- s_data[!is.na(s_data[[1]]), ]
    if(nrow(s_data) == 0) return(NULL)
    
    # Reshape
    colnames(s_data) <- c("student_number", paste0("idx_", 1:length(meta$attr)))
    s_data %>%
      pivot_longer(-student_number, names_to = "idx", values_to = "score") %>%
      mutate(
        i = as.numeric(gsub("idx_", "", idx)),
        course_name = as.character(raw[1, 1]),
        attribute = str_extract(meta$attr[i], "\\d+"),
        indicator = str_extract(meta$ind[i], "\\d+\\.?\\d*"),
        level_assessed = meta$lvl[i],
        student_id = paste(course_name, student_number, sep = "_")
      ) %>%
      filter(!is.na(score)) %>%
      select(student_id, course_name, attribute, indicator, level_assessed, score)
    
  }, error = function(e) {
    log_event(e$message, basename(fp), "ERROR")
    return(NULL) 
  })
}

cat("Reading files...\n")
all_data <- bind_rows(lapply(file_paths, process_file))

if (nrow(all_data) == 0) stop("No valid data found.")

  # Global Formatting
all_data <- all_data %>%
  mutate(
    course_name = factor(course_name),
    attribute = factor(attribute, levels = sort(unique(as.numeric(attribute)))),
    indicator = factor(indicator, levels = sort(unique(as.numeric(indicator)))),
    level_assessed = factor(level_assessed, levels = c("I", "D", "A")),
    score_label = factor(score, levels = c("4","3","2","1"), labels = c("4 - Fail","3 - Minimal","2 - Adequate","1 - Exceeds"))
  )

# Output Directory
out_dir <- file.path(dirname(file_paths[1]), paste0(parent_folder, "_grade_outputs"))
dir.create(out_dir, showWarnings = FALSE)
write.csv(all_data, file.path(out_dir, paste0(parent_folder, "_cleaned_grade_data.csv")), row.names = FALSE)

# 4. Processing Loops (Amalgamated & Courses)
# -------------------------------------------

# We treat "Amalgamated" as just another item in a list of things to process
# This unifies the logic completely.
processing_queue <- list(list(name = "Amalgamated", data = all_data, folder = out_dir))

# Add each course to the queue
for (crs in unique(all_data$course_name)) {
clean_n <- extract_course_code(as.character(crs))
  c_folder <- file.path(out_dir, paste0(parent_folder, "_", clean_n))
  dir.create(c_folder, showWarnings = FALSE)
  
  processing_queue[[length(processing_queue) + 1]] <- list(
    name = as.character(crs),
    data = all_data %>% filter(course_name == crs),
    folder = c_folder,
    prefix = paste0(parent_folder, "_", clean_n) # Prefix for files
  )
}

cat("Generating outputs...\n")

# Process the Queue
for (item in processing_queue) {
  # Warning Handler Context
  current_context <- paste("Processing", item$name)
  
  withCallingHandlers({
    
    file_pre <- if(item$name == "Amalgamated") paste0(parent_folder, "_amalgamated") else item$prefix
    
    generate_output_set(
      data = item$data,
      folder_path = item$folder,
      file_prefix = file_pre,
      title_prefix = item$name,
      context_str = current_context,
      footnote = paste("Source:", parent_folder)
    )
    
  }, warning = function(w) {
    # 1. Log the warning with specific location context
    log_event(w$message, current_context, "WARNING")
    # 2. Suppress from console
    invokeRestart("muffleWarning")
  }, error = function(e) {
    log_event(e$message, current_context, "ERROR")
  })
}

# 5. Isolated Plots (Optional: Kept separate as logic differs slightly)
# ---------------------------------------------------------------------
cat("Generating isolated plots...\n")
iso_dir <- file.path(out_dir, paste0(parent_folder, "_isolated_amalgamated_plots"))
vars <- c("attribute", "indicator", "level_assessed")

withCallingHandlers({
  for (v in vars) {
    sub_dir <- file.path(iso_dir, paste0(v, "_plots"))
    dir.create(sub_dir, recursive = TRUE, showWarnings = FALSE)
    
    for (val in na.omit(unique(all_data[[v]]))) {
      # Trick: use generic plotter but filter data first
      sub_d <- all_data %>% filter(.data[[v]] == val)
      # Create dummy column for x-axis so bar isn't huge
      sub_d$dummy <- "" 
      
      p <- plot_distribution(sub_d, "dummy", paste(v, val), paste("Source:", parent_folder)) + 
        theme(axis.text.x = element_blank(), axis.ticks.x = element_blank()) + labs(x = NULL)
      
tryCatch({
        ggsave(file.path(sub_dir, paste0(parent_folder, "_", v, "_", gsub("\\.", "_", val), ".png")), 
               p, width = 4, height = 6, dpi = 300)
      }, error = function(e) {
        log_event(paste("SKIPPED - ggsave failed:", e$message), 
                  paste("Isolated", v, val), "ERROR")
      })
    }
    
    # For attribute and indicator, create additional subfolder with level separation
    if (v %in% c("attribute", "indicator")) {
      level_sub_dir <- file.path(iso_dir, paste0(v, "_by_level_plots"))
      dir.create(level_sub_dir, recursive = TRUE, showWarnings = FALSE)
      
      for (val in na.omit(unique(all_data[[v]]))) {
        sub_d <- all_data %>% filter(.data[[v]] == val)
        
        # Use level_assessed as x-axis
        p_level <- plot_distribution(sub_d, "level_assessed", paste(v, val), paste("Source:", parent_folder))
        
tryCatch({
          ggsave(file.path(level_sub_dir, paste0(parent_folder, "_", v, "_", gsub("\\.", "_", val), "_by_level.png")), 
                 p_level, width = 6, height = 6, dpi = 300)
        }, error = function(e) {
          log_event(paste("SKIPPED - ggsave failed:", e$message), 
                    paste("Isolated by level", v, val), "ERROR")
        })
      }
    }
  }
}, warning = function(w) {
  log_event(w$message, "Isolated Plots", "WARNING")
  invokeRestart("muffleWarning")
})

cat(paste("\n✓ Done! Check log at:", log_file, "\n"))



