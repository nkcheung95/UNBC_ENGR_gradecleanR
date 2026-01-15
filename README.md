# 📊 UNBC_ENGR_gradecleanR

**A specialized grade cleaning and visualization tool developed for the UNBC Engineering Department.**

This tool automates the processing of grade data into standardized visual reports. It handles package dependencies, data cleaning, and figure generation (Indicators, Attributes, and Levels) across multiple course files simultaneously.

---

## 🚀 Quick Start

To run the most recent version of the script immediately, paste the following command into your **RStudio Console**:

```r
source("https://github.com/nkcheung95/UNBC_ENGR_gradecleanR/blob/main/UNBC_ENGR_grade_summarizer.r?raw=TRUE")

```

---

## 🛠 Prerequisites

Before running the script, ensure you have the following installed:

1. **Base R**: [Download from CRAN](https://cran.r-project.org/bin/windows/base/)
2. **RStudio Desktop**: [Download from Posit](https://posit.co/download/rstudio-desktop/)

---

# How to Use

## Steps

1. **Run the Script** – Copy and paste the Quick Start command into RStudio.
2. **Automatic Setup** – Required packages install automatically if missing.
3. **Select Files** – Choose all Excel files from the dialog. Files must be in the same folder.
4. **Results** – Outputs save to `[source]_out` in your input directory.
5. **Log File** – A timestamped log (e.g., `[source]_YYYYMMDD_HHMMSS_log.txt`) is generated for each run.

---

## Troubleshooting

**File dialog hidden?** Minimize windows to find it—it may open behind RStudio.

**Errors?** Check the log file, then report issues on GitHub with:
- Error description
- Relevant log excerpts  
- R version and OS

---

## Best Practice

Use `source()` instead of local copies to get the latest updates.
