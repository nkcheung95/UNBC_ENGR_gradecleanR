# Grade Summary Data Cleaning Script (Optimized)
# ===============================================

# 1. Setup and Helper Functions
# ---------------------------------
required_packages <- c("ggtext", "ggplot2", "readxl", "dplyr", "tidyr", "tcltk", "stringr", "openxlsx", "rlang", "gtools", "grid", "rsvg", "png")
for (pkg in required_packages) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  }
}

# --- Logo Setup from URL ---
# Replace this URL with your logo's direct URL
LOGO_URL <- "https://www.unbc.ca/themes/custom/unbc/gfx/logo-unbc.svg"
# Download and load SVG logo
logo_img <- NULL
tryCatch({
  temp_svg <- tempfile(fileext = ".svg")
  temp_png <- tempfile(fileext = ".png")
  download.file(LOGO_URL, temp_svg, mode = "wb", quiet = TRUE)
  # Convert SVG to PNG raster for plotting
  rsvg::rsvg_png(temp_svg, temp_png, width = 1200)
  logo_img <- readPNG(temp_png)
  cat("Logo loaded successfully from URL!\n")
}, error = function(e) {
  cat("Error loading logo from URL:", e$message, "\nProceeding without logo.\n")
})
# --- Logger Setup ---
log_event <- function(msg, context = "General", type = "INFO") {
  entry <- sprintf("[%s] [%s] {%s}: %s", Sys.time(), type, context, msg)
  if (exists("log_file")) write(entry, file = log_file, append = TRUE)
  if (type != "WARNING") cat(paste0(entry, "\n"))
}
# --- Header Text Setup ---
HEADER_TEXT <- "UNBC School of Engineering"  # Edit this text as needed
HEADER_SUBTITLE <- "Graduate Attribute Assessment 2025-2026"  # Optional subtitle
# --- Generic Clean Filename (10 char limit for source folder) ---
clean_filename <- function(name, max_len = 100) {
  name %>% 
    gsub("[^A-Za-z0-9_-]", "_", .) %>%
    gsub("_{2,}", "_", .) %>% 
    gsub("^_|_$", "", .) %>%
    substr(1, max_len)
}

# --- Extract Course Code (first 3 digits, limit to 10 chars) ---
extract_course_code <- function(course_name) {
  match <- str_extract(course_name, "^.*?\\d{3}")
  if (is.na(match)) return(clean_filename(course_name, 10))
  return(clean_filename(match, 10))
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

# --- Generic Plotter with Logo Header ---
plot_distribution <- function(data, group_col, title, footnote = NULL) {
  grp_sym <- sym(group_col)
  
  stats <- data %>%
    group_by(!!grp_sym) %>%
    summarise(
      label = sprintf("Meet: %.1f%%\nmean: %.1f\nSD: %.2f", 
                      sum(score %in% 1:3)/n()*100, mean(as.numeric(score)), sd(as.numeric(score))),
      .groups = 'drop'
    )
  
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
    theme(
      axis.text.x = element_text(angle = 45, hjust = 1), 
      plot.title = element_text(hjust = 0.5, face = "bold"),
      plot.margin = margin(t = 18, r = 10, b = 10, l = 10)  # ~50 pixels at 300 dpi
    )
  if (!is.null(footnote)) {
    p <- p + labs(caption = footnote) + 
      theme(plot.caption = element_text(hjust = 0.5, size = 9, color = "gray30"),
            plot.margin = margin(t = 18, r = 10, b = 10, l = 10))
  }
  
  return(p)
}

# --- Save plot with logo header ---
save_plot_with_logo <- function(plot, filepath, width = 8, height = 6, dpi = 300) {
  tryCatch({
    if (!is.null(logo_img) || exists("HEADER_TEXT")) {
      # Create composite with logo and/or text header
      png(filepath, width = width * dpi, height = (height + 0.8) * dpi, res = dpi)
      
      # Layout: header at top, plot below (no gap)
      grid.newpage()
      pushViewport(viewport(layout = grid.layout(2, 1, heights = unit(c(0.8, height), "inches"))))
      
      # Add header (logo + text)
      pushViewport(viewport(layout.pos.row = 1, layout.pos.col = 1))
      
      # Add green background
      grid.rect(gp = gpar(fill = "#035642", col = NA))
      
      if (!is.null(logo_img)) {
        grid.raster(logo_img, x = 0.05, y = 0.5, width = unit(1.2, "inches"), just = "left")
      }
      
      if (exists("HEADER_TEXT")) {
        text_x <- if (!is.null(logo_img)) 0.30 else 0.5
        text_just <- if (!is.null(logo_img)) "left" else "center"
        
        grid.text(HEADER_TEXT, x = text_x, y = 0.65, just = text_just,
                  gp = gpar(fontsize = 16, fontface = "bold", col = "white"))
        
        if (exists("HEADER_SUBTITLE") && nchar(HEADER_SUBTITLE) > 0) {
          grid.text(HEADER_SUBTITLE, x = text_x, y = 0.35, just = text_just,
                    gp = gpar(fontsize = 12, col = "white"))
        }
      }
      
      popViewport()
      
      # Add main plot (directly below header, no gap)
      pushViewport(viewport(layout.pos.row = 2, layout.pos.col = 1))
      plot_grob <- ggplotGrob(plot)
      grid.draw(plot_grob)
      popViewport()
      
      dev.off()
    } else {
      # No logo or header text, save normally
      ggsave(filepath, plot, width = width, height = height, dpi = dpi)
    }
  }, error = function(e) {
    stop(paste("Plot save failed:", e$message))
  })
}
# --- Master Output Generator ---
generate_output_set <- function(data, folder_path, file_prefix, title_prefix, context_str, footnote = NULL) {
  
  wb <- createWorkbook()
  vars <- c("attribute", "indicator", "level_assessed")
  
  for (var in vars) {
    # Shortened sheet names
    sheet_name <- case_when(
      var == "attribute" ~ "By_Attr",
      var == "indicator" ~ "By_Ind",
      var == "level_assessed" ~ "By_Lvl"
    )
    
    addWorksheet(wb, sheet_name)
    writeData(wb, sheet_name, calc_stats(data, var))
    
    # Shortened figure names: attr, ind, lvl
    var_short <- case_when(
      var == "attribute" ~ "attr",
      var == "indicator" ~ "ind",
      var == "level_assessed" ~ "lvl"
    )
    
    n_items <- length(unique(data[[var]]))
    p_width <- if(var == "level_assessed") max(6, n_items * 2) else max(8, n_items * 0.8)
    
    p <- plot_distribution(data, var, paste(title_prefix, "-", var), footnote)
    tryCatch({
      save_plot_with_logo(p, file.path(folder_path, paste0(file_prefix, "_fig_", var_short, ".png")), 
                          width = p_width, height = 6, dpi = 300)
    }, error = function(e) {
      log_event(paste("SKIPPED - plot save failed:", e$message), 
                paste(context_str, "- figure", var), "ERROR")
    })
  }
  
  saveWorkbook(wb, file.path(folder_path, paste0(file_prefix, "_sum.xlsx")), overwrite = TRUE)
  log_event(paste("Saved outputs to", folder_path), context_str)
}

# 2. File Selection
# -----------------
cat("\n>>> PLEASE SELECT FILES IN THE POP-UP WINDOW <<<\n")
flush.console()

file_paths <- tk_choose.files(caption = "Select Excel Files", filters = matrix(c("Excel", ".xlsx;.xls"), 1, 2))
if (length(file_paths) == 0) stop("No files selected.")

# Limit parent folder to 10 chars
parent_folder <- substr(clean_filename(basename(dirname(file_paths[1]))), 1, 10)

# Init Log with timestamp
timestamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
log_file <- file.path(dirname(file_paths[1]), paste0(parent_folder, "_", timestamp, "_log.txt"))
write(paste("Run Date:", Sys.time(), "\nSource:", parent_folder, "\n---"), file = log_file)

# 3. Data Ingestion
# -----------------
process_file <- function(fp) {
  tryCatch({
    raw <- read_excel(fp, sheet = 1, col_names = FALSE)
    if(nrow(raw) < 11) return(NULL)
    
    meta_rows <- list(attr = 6, ind = 7, lvl = 8)
    meta <- lapply(meta_rows, function(r) as.character(raw[r, -1]))
    
    s_data <- raw[11:nrow(raw), ]
    s_data <- s_data[!is.na(s_data[[1]]), ]
    if(nrow(s_data) == 0) return(NULL)
    
    colnames(s_data) <- c("student_number", paste0("idx_", 1:length(meta$attr)))
    s_data %>%
      pivot_longer(-student_number, names_to = "idx", values_to = "score") %>%
      mutate(
        i = as.numeric(gsub("idx_", "", idx)),
        course_name = as.character(raw[1, 1]),
        attribute = str_extract(meta$attr[i], "\\d+"),
        indicator = as.character(round(as.numeric(str_extract(meta$ind[i], "\\d+\\.?\\d*")), 1)),
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

all_data <- all_data %>%
  mutate(
    course_name = factor(course_name),
    attribute = factor(attribute, levels = sort(unique(as.numeric(attribute)))),
    indicator = factor(indicator, levels = gtools::mixedsort(unique(indicator))),
    level_assessed = factor(level_assessed, levels = c("I", "D", "A")),
    score_label = factor(score, levels = c("4","3","2","1"), labels = c("4 - Fail","3 - Minimal","2 - Adequate","1 - Exceeds"))
  )

# Output Directory (shortened to "out")
out_dir <- file.path(dirname(file_paths[1]), paste0(parent_folder, "_out"))
dir.create(out_dir, showWarnings = FALSE)
write.csv(all_data, file.path(out_dir, paste0(parent_folder, "_clean.csv")), row.names = FALSE)

# 4. Processing Loops
# -------------------------------------------
processing_queue <- list(list(name = "Amalg", data = all_data, folder = out_dir))

for (crs in unique(all_data$course_name)) {
  clean_n <- extract_course_code(as.character(crs))
  c_folder <- file.path(out_dir, paste0(parent_folder, "_", clean_n))
  dir.create(c_folder, showWarnings = FALSE)
  
  processing_queue[[length(processing_queue) + 1]] <- list(
    name = as.character(crs),
    data = all_data %>% filter(course_name == crs),
    folder = c_folder,
    prefix = paste0(parent_folder, "_", clean_n)
  )
}

cat("Generating outputs...\n")

for (item in processing_queue) {
  current_context <- paste("Processing", item$name)
  
  withCallingHandlers({
    
    file_pre <- if(item$name == "Amalg") paste0(parent_folder, "_amalg") else item$prefix
    
    generate_output_set(
      data = item$data,
      folder_path = item$folder,
      file_prefix = file_pre,
      title_prefix = item$name,
      context_str = current_context,
      footnote = paste("Source:", parent_folder)
    )
    
  }, warning = function(w) {
    log_event(w$message, current_context, "WARNING")
    invokeRestart("muffleWarning")
  }, error = function(e) {
    log_event(e$message, current_context, "ERROR")
  })
}

# 5. Isolated Plots
# ---------------------------------------------------------------------
cat("Generating isolated plots...\n")
iso_dir <- file.path(out_dir, paste0(parent_folder, "_iso"))
vars <- c("attribute", "indicator", "level_assessed")

withCallingHandlers({
  for (v in vars) {
    # Shortened subfolder names: attr_plt, ind_plt, lvl_plt
    var_short <- case_when(
      v == "attribute" ~ "attr",
      v == "indicator" ~ "ind",
      v == "level_assessed" ~ "lvl"
    )
    
    sub_dir <- file.path(iso_dir, paste0(var_short, "_plt"))
    dir.create(sub_dir, recursive = TRUE, showWarnings = FALSE)
    
    for (val in na.omit(unique(all_data[[v]]))) {
      sub_d <- all_data %>% filter(.data[[v]] == val)
      sub_d$dummy <- "" 
      
      p <- plot_distribution(sub_d, "dummy", paste(v, val), paste("Source:", parent_folder)) + 
        theme(axis.text.x = element_blank(), axis.ticks.x = element_blank()) + labs(x = NULL)
      
      tryCatch({
        save_plot_with_logo(p, file.path(sub_dir, paste0(parent_folder, "_", var_short, "_", gsub("\\.", "_", val), ".png")), 
                            width = 4, height = 6, dpi = 300)
      }, error = function(e) {
        log_event(paste("SKIPPED - plot save failed:", e$message), 
                  paste("Isolated", v, val), "ERROR")
      })
    }
    
    # For attribute and indicator, create level separation plots
    if (v %in% c("attribute", "indicator")) {
      level_sub_dir <- file.path(iso_dir, paste0(var_short, "_lv"))
      dir.create(level_sub_dir, recursive = TRUE, showWarnings = FALSE)
      
      for (val in na.omit(unique(all_data[[v]]))) {
        sub_d <- all_data %>% filter(.data[[v]] == val)
        
        p_level <- plot_distribution(sub_d, "level_assessed", paste(v, val), paste("Source:", parent_folder))
        
        tryCatch({
          save_plot_with_logo(p_level, file.path(level_sub_dir, paste0(parent_folder, "_", var_short, "_", gsub("\\.", "_", val), "_lv.png")), 
                              width = 6, height = 6, dpi = 300)
        }, error = function(e) {
          log_event(paste("SKIPPED - plot save failed:", e$message), 
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
