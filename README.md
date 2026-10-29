# Student Performance Data Analysis: Exploratory Data Analysis & Multivariate Visualization

[![Python Version](https://img.shields.io/badge/Python-3.10%20%7C%203.11%20%7C%203.12-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-2.2+-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![NumPy](https://img.shields.io/badge/NumPy-1.26+-013243?logo=numpy&logoColor=white)](https://numpy.org/)
[![Seaborn](https://img.shields.io/badge/Seaborn-0.13+-3776AB)](https://seaborn.pydata.org/)
[![Matplotlib](https://img.shields.io/badge/Matplotlib-3.9+-11557c)](https://matplotlib.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 📌 Academic Project Metadata

| Field | Institutional Details |
| :--- | :--- |
| **Student Name** | **Mithun** |
| **University UID** | **25BCD10179** |
| **Degree / Program** | **Bachelor of Computer Applications — BCA (DATA SCIENCE)** |
| **Academic Term** | **Semester: 3 \| Section: 25BCD-2-B** |
| **Course Title & Code** | **Python Programming & Data Science (`25CAT-204`)** |
| **Academic Year** | **2026 – 2027** |
| **Institution** | **Chandigarh University, Gharuan, Mohali, Punjab** |

---

## 📖 Executive Summary & Problem Formulation

Educational Data Mining (EDM) and Learning Analytics (LA) empower academic institutions to transition from retrospective post-examination reviews to proactive, predictive student interventions. This project implements an end-to-end, reproducible **Python Data Science Pipeline** to analyze summative assessment performance, behavioral indicators, and socio-demographic factors across a cohort of **1,000 undergraduate students**.

The analytical framework addresses five foundational educational questions:
1. **Study Habit Elasticity:** What is the quantitative rate of return of weekly self-study duration on examination marks ($r = +0.542$)?
2. **Attendance Thresholds:** How strongly does physical classroom attendance correlate with overall score ($r = +0.241$), and what minimum threshold protects against academic failure?
3. **Curricular Gender Parity:** Do male and female students demonstrate statistically significant differences between STEM disciplines (Mathematics) and verbal disciplines (Reading and Writing)?
4. **Intergenerational Educational Capital:** To what extent does parental educational attainment influence academic achievement and compress score variance?
5. **Institutional Coaching Efficacy:** What is the net value-add of structured test preparation programs (+8.03 point boost)?

---

## 🏛️ Pipeline Architecture

The system is architected as a modular, five-stage deterministic pipeline following software engineering best practices:

```
+-----------------------------------------------------------------------------------------+
|                    STAGE 1: SYNTHETIC COHORT GENERATION (generate_dataset.py)            |
|  NumPy Random Distributions -> Demographic & Behavioral Matrix -> student_performance.csv |
+--------------------------------------------+--------------------------------------------+
                                             |
                                             v
+-----------------------------------------------------------------------------------------+
|                    STAGE 2: DATA CLEANSING & PREPROCESSING (Pandas)                     |
|  Duplicate Detection -> Null Value Audit -> Median Imputation -> Text Standardization   |
+--------------------------------------------+--------------------------------------------+
                                             |
                                             v
+-----------------------------------------------------------------------------------------+
|                    STAGE 3: FEATURE ENGINEERING & STATISTICAL AGGREGATION               |
|  Composite Scores (total_score) -> Mean Average -> Letter Grades (A-F) -> Groupby Stats  |
+--------------------------------------------+--------------------------------------------+
                                             |
                                             v
+-----------------------------------------------------------------------------------------+
|                    STAGE 4: VISUAL ANALYTICS & GRAPHICS ENGINE                          |
|  Matplotlib (Line, Bar, Hist, Scatter) + Seaborn (Box, Violin, Heatmap, Count, Pairplot)|
+--------------------------------------------+--------------------------------------------+
                                             |
                                             v
+-----------------------------------------------------------------------------------------+
|                    STAGE 5: AUTOMATED REPORT GENERATION (build_documents.py)            |
|  python-docx Engine -> Case Study Report (.docx) + Comprehensive Mini Project (.docx)   |
+-----------------------------------------------------------------------------------------+
```

---

## 📊 Dataset Schema & Preprocessing Methodology

The dataset encompasses 10 primary attributes and 3 engineered features across 1,000 unique records:

| Attribute | Data Type | Scale | Range | Description |
| :--- | :--- | :--- | :--- | :--- |
| `student_id` | String | Nominal | STD-1001 – STD-2000 | Unique student registration roll code |
| `gender` | String | Nominal | `female`, `male` | Self-identified student gender |
| `parental_education` | String | Ordinal | 5 degree tiers | Highest degree completed by parents/guardians |
| `study_hours_per_week` | Float | Ratio | 5.0 – 39.0 hrs | Self-directed weekly study duration outside class |
| `attendance_percentage`| Float | Ratio | 50.0% – 99.0% | Percentage of semester classes attended |
| `test_prep_course` | String | Binary | `none`, `completed` | Enrollment in structured preparatory workshop |
| `math_score` | Float | Interval | 60.0 – 100.0 | Standardized assessment mark in Mathematics |
| `reading_score` | Float | Interval | 65.0 – 100.0 | Standardized assessment mark in Reading Comprehension |
| `writing_score` | Float | Interval | 64.0 – 100.0 | Standardized assessment mark in Written Expression |
| `total_score` | Float | Derived | 195.0 – 300.0 | Sum of math, reading, and writing marks |
| `average_score` | Float | Derived | 65.0% – 100.0% | Calculated percentage mean across all 3 subjects |
| `grade` | String | Derived | `A`, `B`, `C`, `D`, `F`| Institutional letter grade classification |

### Data Cleansing Highlights
- **Deduplication:** Exactly 3 synthetic duplicate records were detected via `df.duplicated()` and eliminated.
- **Robust Median Imputation:** Missing entries in `math_score` (5 nulls), `study_hours_per_week` (4 nulls), and `attendance_percentage` (3 nulls) were imputed using median values rather than mean statistics to provide resistance against distributional skewness and outliers.
- **String Standardization:** Applied vectorized `.str.strip().str.lower()` to eliminate whitespace inconsistencies.

---

## 📈 Key Empirical Findings

1. **Study Hours as Primary Determinant:** Weekly self-directed study duration displays the highest linear correlation with average score ($r = +0.542, p < 0.0001$). Optimal performance gains occur between 18 and 24 hours/week; beyond 28 hours, diminishing marginal returns set in.
2. **Test Preparation Dividend:** Students who completed the preparatory course achieved a mean composite average of **93.85%**, compared to **85.82%** for non-completers—an empirical boost of **+8.03 percentage points**.
3. **Disciplinary Gender Strengths:** While overall performance is comparable, female students outperform males in verbal disciplines (**Reading:** 91.23 vs 87.70, **Writing:** 91.42 vs 87.21). Male students maintain a modest edge in **Mathematics** (88.61 vs 86.00).
4. **Literacy Interdependence:** Reading and Writing scores exhibit an extraordinary Pearson correlation coefficient of **$r = +0.884$**, confirming shared cognitive fundamentals.
5. **Socio-Demographic Gradient:** Children of Master's degree holders average **94.44%**, whereas children of high school diploma holders average **83.21%**—an attainment divide of **11.23 percentage points**.
6. **Attendance Risk Filter:** Attendance percentage ($r = +0.241$) serves as a protective barrier; students with attendance below 65% represent 88% of all lower-tier grades (Grade C and D).

---

## 🖼️ Visual Analytics Gallery (Figures 1 to 9)

All figures were generated at **300 DPI** using Matplotlib and Seaborn and are available in the `figures/` directory:

| Figure ID | Chart Type | Dimension Examined | Core Insight |
| :--- | :--- | :--- | :--- |
| **Figure 1** | Matplotlib Line Chart | Study Hours vs. Score | Non-linear scaling with plateauing beyond 28 hrs/week |
| **Figure 2** | Matplotlib Grouped Bar | Subject Marks by Gender | Female verbal leadership vs. male math advantage |
| **Figure 3** | Matplotlib Hist + KDE | Overall Score Distribution | Gaussian bell curve ($\mu = 88.70, s = 6.25$) |
| **Figure 4** | Matplotlib Scatter + OLS | Attendance vs. Score | Positive regression slope ($+1.52$ pts per 10% attendance) |
| **Figure 5** | Seaborn Box Plot | Score by Parental Education | Monotonic progression from 83.21 to 94.44 |
| **Figure 6** | Seaborn Violin Plot | Math Density by Gender | Bimodal upper lobe in male mathematics scores |
| **Figure 7** | Seaborn Heatmap | Pearson Correlation Matrix | Annotated matrix showing Reading-Writing collinearity ($0.884$) |
| **Figure 8** | Seaborn Count Plot | Test Prep by Gender | Equal participation with +8.03 point boost |
| **Figure 9** | Seaborn Pair Plot | Multivariate Scatter & KDE | 3x3 matrix showing dense collinear point clouds |

---

## 📁 Repository Directory Structure

```plaintext
python-student-performance/
│
├── data/
│   └── student_performance.csv               # 1,000 student analytical dataset
│
├── docs/
│   ├── Case_Study_Student_Performance.docx   # Strict 7-section Case Study Report
│   └── Mini_Project_Report_Student_Performance.docx # Comprehensive ~28-30 Page Report
│
├── figures/
│   ├── fig1_line_study_vs_marks.png          # Weekly study hours line chart
│   ├── fig2_bar_subject_by_gender.png        # Grouped bar chart
│   ├── fig3_hist_score_distribution.png      # Histogram with KDE curve
│   ├── fig4_scatter_attendance_vs_marks.png  # Scatter plot with regression line
│   ├── fig5_box_scores_by_parent_education.png # Box plot across education tiers
│   ├── fig6_violin_math_by_gender.png        # Violin probability density
│   ├── fig7_heatmap_correlation.png          # Correlation matrix heatmap
│   ├── fig8_count_test_prep.png              # Test prep count plot
│   └── fig9_pairplot_scores.png              # Multivariate pair plot
│
├── src/
│   ├── generate_dataset.py                   # Synthetic cohort generator (NumPy/Pandas)
│   ├── student_analysis.py                   # Data cleaning, EDA & visualization engine
│   ├── report_utils.py                       # Document formatting & table styling helper
│   └── build_documents.py                    # Automated Word (.docx) document compiler
│
├── .gitignore                                # Python & environment exclusions
├── requirements.txt                          # Pip dependency specifications
└── README.md                                 # Comprehensive project documentation
```

---

## 🚀 Installation & Execution Guide

### 1. Clone the Repository
```bash
git clone <repository_url>
cd python-student-performance
```

### 2. Set Up Virtual Environment (Optional but Recommended)
```bash
python -m venv venv
# On Windows:
venv\Scripts\activate
# On Linux/macOS:
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the Pipeline
```bash
# Step 1: Synthesize dataset
python src/generate_dataset.py

# Step 2: Execute data cleaning, EDA & generate all 9 visual figures
python src/student_analysis.py

# Step 3: Compile both Word reports
python src/build_documents.py
```

---

## 📄 Formal Deliverables

1. **`Case_Study_Student_Performance.docx`**: Strictly adheres to the 7-section structure defined on pages 3–4 of the course guidelines:
   - *1. Problem Statement | 2. Dataset Description | 3. Data Cleaning Using Pandas | 4. Data Analysis Using NumPy & Pandas | 5. Data Visualization (9 Figures) | 6. Findings (7 Key Observations) | 7. Conclusion*
2. **`Mini_Project_Report_Student_Performance.docx`**: Comprehensive **28–30 page** academic mini project report featuring:
   - *Preliminaries: Chandigarh University banner, Student block (Mithun - 25BCD10179), Abstract, Table of Contents, List of Figures, List of Tables*
   - *8 Detailed Chapters with complete theoretical foundations, pipeline architectures, commented source code, 11 styled tables, and deep pedagogical evaluations of all 9 figures.*

---

## 📜 License & Acknowledgements

Developed as part of the academic curriculum for **Python Programming & Data Science (`25CAT-204`)** at **Chandigarh University**.  
Licensed under the [MIT License](LICENSE).
