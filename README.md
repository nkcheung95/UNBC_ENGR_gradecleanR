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

## 📖 How to Use

1. **Execute the Script**: Copy and paste the command in the QUICK START section into RStudio.
2. **Automatic Setup**: The program will automatically check for and download any required packages (such as `ggplot2`, `dplyr`, and `ggtext`).
3. **Select Files**:
* A file selection window will appear.
* Navigate to your data folder and select all the Excel files you wish to process.
* **Note:** Files must be in the same folder to be selected together.


4. **Retrieve Results**:
* Once processing is complete, look in the directory where your source files were located.
* A new folder named `grade_outputs` will have been created containing all your generated figures and summaries.



---

## ⚠️ Important Note on File Selection

The file selection dialog box often opens **behind** RStudio or other active windows. If the program seems to be "hanging" after you run the command:

* Minimize RStudio and other open windows.
* Look for the "Select Excel files with grade summaries" popup.


### 💡 Pro-Tip

*Always use the `source()` command provided above rather than a local copy to ensure you are benefiting from the latest bug fixes and feature updates.*
