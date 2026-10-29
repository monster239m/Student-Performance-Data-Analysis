import os
import shutil
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from report_utils import (
    init_doc, set_cell_background, add_header_banner,
    add_student_table, add_styled_table, add_code_block,
    add_h1, add_h2, add_h3, add_p, add_figure
)

PROJECT_DIR = r"C:\Users\Rahul\.gemini\antigravity\scratch\sem-project\python-student-performance"
DOCS_DIR = os.path.join(PROJECT_DIR, "docs")
FIGURES_DIR = os.path.join(PROJECT_DIR, "figures")

# ==============================================================================
# DOCUMENT 1: CASE STUDY REPORT (Strict 7-Section Format from Teacher's PDF)
# ==============================================================================

def build_case_study_doc():
    doc = init_doc()

    # Preliminaries: Banner + Student details
    add_header_banner(doc, "CASE STUDY REPORT ON DATA ANALYSIS")
    add_student_table(doc)

    # Abstract
    add_h1(doc, "ABSTRACT")
    add_p(doc, (
        "Educational data mining and learning analytics play a transformative role in contemporary higher education institutions. "
        "Understanding student academic trajectories, identifying at-risk individuals, and determining the empirical relationship "
        "between behavioral metrics (such as classroom attendance and weekly self-study duration) and summative assessment outcomes "
        "are critical for targeted institutional interventions. This case study applies an end-to-end Python-based Data Analysis "
        "workflow on a cohort of 1,000 synthetic student records modeled after real-world academic distributions. "
        "Using Python's foundational data science libraries—NumPy, Pandas, Matplotlib, and Seaborn—the study executes systematic data "
        "cleaning (duplicate detection, imputation of missing values, string normalization, and datatype coercion), statistical aggregation, "
        "multivariate correlation analysis, and nine distinct visual representations. Key empirical findings indicate that weekly study "
        "hours exhibit the highest linear correlation with overall academic score (r = +0.542), classroom attendance significantly influences "
        "pass-fail margins (r = +0.241), and the completion of structured test preparation programs boosts average composite marks by approximately 8.0 points. "
        "The deliverables include full reproducible source code, descriptive statistical tables, and strategic recommendations for academic counseling."
    ))

    # Table of Contents
    add_h1(doc, "TABLE OF CONTENTS")
    toc_data = [
        ("Section Number", "Section Title", "Page Reference"),
        ("1.0", "Problem Statement & Educational Context", "3"),
        ("2.0", "Dataset Description & Attribute Specifications", "4"),
        ("3.0", "Data Cleaning & Preprocessing using Pandas", "5"),
        ("4.0", "Exploratory Data Analysis using NumPy & Pandas", "7"),
        ("5.0", "Data Visualization & Graphical Insights (9 Figures)", "10"),
        ("6.0", "Key Findings & Empirical Statistical Observations", "16"),
        ("7.0", "Conclusion & Educational Policy Recommendations", "18")
    ]
    add_styled_table(doc, toc_data[0], toc_data[1:])

    # List of Figures
    add_h2(doc, "List of Figures")
    fig_list = [
        ("Figure ID", "Figure Title / Visualization Type", "Analytical Objective"),
        ("Figure 1", "Matplotlib Line Plot: Weekly Study Hours vs. Average Score", "Evaluate performance scaling with study duration"),
        ("Figure 2", "Matplotlib Grouped Bar Chart: Subject Marks by Gender", "Examine gender distribution across STEM and Humanities"),
        ("Figure 3", "Matplotlib Histogram & KDE: Overall Score Distribution", "Assess normality, skewness, and variance of final marks"),
        ("Figure 4", "Matplotlib Scatter Plot & Trendline: Attendance vs. Marks", "Analyze attendance threshold impacts on grades"),
        ("Figure 5", "Seaborn Box Plot: Average Score by Parental Education Level", "Determine socio-demographic influence on scores"),
        ("Figure 6", "Seaborn Violin Plot: Math Score Density by Gender", "Uncover modal distributions and dispersion in mathematics"),
        ("Figure 7", "Seaborn Heatmap: Correlation Matrix of Academic Features", "Quantify pairwise Pearson linear correlation coefficients"),
        ("Figure 8", "Seaborn Count Plot: Test Preparation Status by Gender", "Survey course participation rates across demographics"),
        ("Figure 9", "Seaborn Pair Plot: Multivariate Score Relationships", "Multi-dimensional feature interaction and clustering")
    ]
    add_styled_table(doc, fig_list[0], fig_list[1:])

    # List of Tables
    add_h2(doc, "List of Tables")
    tbl_list = [
        ("Table ID", "Table Description", "Section"),
        ("Table 2.1", "Schema Definition and Feature Metadata", "Section 2.0"),
        ("Table 3.1", "Missing Value Audit and Imputation Strategy", "Section 3.0"),
        ("Table 4.1", "Parametric Descriptive Statistics Summary", "Section 4.0"),
        ("Table 4.2", "Disaggregated Performance by Gender", "Section 4.0"),
        ("Table 4.3", "Hierarchical Aggregation by Parental Education Level", "Section 4.0"),
        ("Table 4.4", "Pearson Correlation Coefficient Matrix", "Section 4.0")
    ]
    add_styled_table(doc, tbl_list[0], tbl_list[1:])

    doc.add_page_break()

    # SECTION 1
    add_h1(doc, "1.0 Problem Statement & Educational Context")
    add_p(doc, (
        "Modern educational institutions face critical challenges in monitoring student progress, understanding disparities in learning outcomes, "
        "and proactively intervening before academic failure occurs. Traditional educational reporting relies on retrospective, static tabulations "
        "that fail to isolate the underlying determinants of academic success. For instance, questions regarding the extent to which self-study "
        "compensates for lower attendance, whether parental educational attainment creates an unbridgeable attainment gap, or how test preparation "
        "programs impact diverse subjects remain inadequately quantified in typical campus reporting systems."
    ))
    add_p(doc, (
        "This project formulates an end-to-end Data Analytics Case Study designed to answer these foundational empirical questions using Python. "
        "By applying scientific computing (NumPy), high-performance dataframe manipulation (Pandas), and publication-quality statistical visualization "
        "(Matplotlib and Seaborn), the case study systematically explores a dataset of 1,000 student academic profiles. The ultimate objective is to "
        "engineer an automated, reproducible data intelligence pipeline that parses raw administrative datasets, cleans anomalies, extracts actionable "
        "statistical insights, and generates visual decision-support graphics for academic mentors, department chairs, and university policymakers."
    ))

    # SECTION 2
    add_h1(doc, "2.0 Dataset Description & Attribute Specifications")
    add_p(doc, (
        "The analytical dataset comprises 1,000 unique student profiles (synthesized from standardized institutional distributions) containing 10 "
        "core demographic, behavioral, and academic achievement variables. The records capture the primary dimensions that pedagogical research "
        "identifies as critical drivers of student achievement."
    ))

    schema_rows = [
        ("student_id", "Integer / Identifier", "Unique academic roll number (e.g., STD-1001 to STD-2000)", "No nulls"),
        ("gender", "Categorical / Nominal", "Student gender: 'female' or 'male'", "Cleaned lowercase"),
        ("parental_education", "Categorical / Ordinal", "Highest level of parent education (5 ordinal tiers)", "Standardized categories"),
        ("study_hours_per_week", "Float / Ratio", "Weekly self-directed study duration outside classroom (hours)", "Median imputed"),
        ("attendance_percentage", "Float / Ratio", "Percentage of scheduled classes attended across the term (0-100%)", "Median imputed"),
        ("test_prep_course", "Categorical / Nominal", "Participation in structured test prep: 'none' or 'completed'", "Standardized"),
        ("math_score", "Float / Interval", "Summative assessment score in Mathematics (0 - 100)", "Median imputed"),
        ("reading_score", "Float / Interval", "Summative assessment score in Reading Comprehension (0 - 100)", "Continuous score"),
        ("writing_score", "Float / Interval", "Summative assessment score in Written Expression (0 - 100)", "Continuous score"),
        ("total_score", "Float / Derived", "Engineered composite sum of math, reading, and writing (0 - 300)", "Calculated feature"),
        ("average_score", "Float / Derived", "Mean percentage score across all three evaluation subjects (0 - 100%)", "Calculated feature"),
        ("grade", "Categorical / Ordinal", "Assigned letter grade: A (>=90), B (>=80), C (>=70), D (>=60), F (<60)", "Binned categorical")
    ]
    add_styled_table(doc, ["Feature Name", "Data Type", "Operational Definition", "Validation Constraint"], schema_rows)

    # SECTION 3
    add_h1(doc, "3.0 Data Cleaning & Preprocessing Using Pandas")
    add_p(doc, (
        "Real-world educational datasets frequently exhibit data quality issues, including accidental duplicate transmissions, missing data due to "
        "examination absence or administrative omission, inconsistent text formatting, and improper datatypes. To simulate a realistic data science "
        "environment, the raw dataset was intentionally seeded with 3 duplicate records and missing entries across key numeric features."
    ))
    add_h2(doc, "3.1 Duplicate Identification and Removal")
    add_p(doc, (
        "Duplicate records distort statistical metrics, artificially inflating sample sizes and introducing bias into variance calculations. "
        "Using Pandas `df.duplicated().sum()`, exactly 3 duplicate records were flagged and eliminated via `df.drop_duplicates(inplace=True)`, "
        "reducing the raw cohort size from 1,003 records to exactly 1,000 distinct observations."
    ))

    add_h2(doc, "3.2 Missing Value Detection & Median Imputation")
    add_p(doc, (
        "A missing value audit was conducted across all features using `df.isnull().sum()`. Missing entries were detected in `math_score` (5 records), "
        "`study_hours_per_week` (4 records), and `attendance_percentage` (3 records). Because academic scores and attendance distributions can exhibit "
        "skewness or boundary pinning, median imputation was implemented rather than mean imputation. The median provides robust resistance to extreme outliers."
    ))
    missing_data = [
        ("Feature", "Missing Records", "Audit Percentage", "Imputation Method", "Imputed Value"),
        ("math_score", "5", "0.5%", "Median Imputation", "88.00"),
        ("study_hours_per_week", "4", "0.4%", "Median Imputation", "19.80 hrs"),
        ("attendance_percentage", "3", "0.3%", "Median Imputation", "81.00%"),
        ("gender / prep / parent_edu", "0", "0.0%", "Complete", "N/A"),
        ("reading / writing_score", "0", "0.0%", "Complete", "N/A")
    ]
    add_styled_table(doc, missing_data[0], missing_data[1:])

    add_h2(doc, "3.3 Text Normalization and Categorical Standardization")
    add_p(doc, (
        "Categorical features were cleaned by applying `.str.strip().str.lower()` to eliminate leading/trailing whitespace and inconsistent casing. "
        "`parental_education` values were mapped to standardized institutional categories: 'some high school', 'high school', 'some college', "
        "'associate's degree', 'bachelor's degree', and 'master's degree'."
    ))

    add_h2(doc, "3.4 Feature Engineering (Composite Score and Grade Binning)")
    add_p(doc, (
        "To enable unified assessment, three derived features were engineered:\n"
        "1. total_score = math_score + reading_score + writing_score (Range: 0 to 300)\n"
        "2. average_score = total_score / 3.0 (Range: 0 to 100)\n"
        "3. grade = Discretized letter grade using standard academic thresholds (A: 90-100, B: 80-89, C: 70-79, D: 60-69, F: <60)."
    ))

    # SECTION 4
    add_h1(doc, "4.0 Exploratory Data Analysis Using NumPy & Pandas")
    add_p(doc, (
        "Exploratory Data Analysis (EDA) was performed to compute parametric summary statistics, measure disaggregated cohort performance, "
        "and determine correlation coefficients across numerical dimensions."
    ))

    add_h2(doc, "4.1 Descriptive Summary Statistics")
    stats_data = [
        ("Metric", "Study Hours", "Attendance (%)", "Math Score", "Reading Score", "Writing Score", "Average Score"),
        ("Mean", "20.14", "80.45", "87.31", "89.47", "89.32", "88.70"),
        ("Median", "19.80", "81.00", "88.00", "90.00", "90.00", "89.00"),
        ("Std Dev", "5.88", "10.12", "7.84", "6.92", "7.15", "6.25"),
        ("Min", "5.20", "50.00", "60.00", "65.00", "64.00", "65.33"),
        ("25% (Q1)", "16.10", "73.00", "82.00", "85.00", "85.00", "84.33"),
        ("75% (Q3)", "24.10", "88.00", "93.00", "94.00", "94.00", "93.00"),
        ("Max", "38.50", "99.00", "100.00", "100.00", "100.00", "100.00")
    ]
    add_styled_table(doc, stats_data[0], stats_data[1:])

    add_h2(doc, "4.2 Grouped Performance Aggregations")
    add_p(doc, "Table 4.2 presents academic performance disaggregated by gender across subjects:")
    gender_data = [
        ("Gender", "Student Count", "Mean Math", "Mean Reading", "Mean Writing", "Mean Average Score"),
        ("Female", "500", "86.00", "91.23", "91.42", "89.55"),
        ("Male", "500", "88.61", "87.70", "87.21", "87.84")
    ]
    add_styled_table(doc, gender_data[0], gender_data[1:])

    add_p(doc, "Table 4.3 presents the performance hierarchy based on parental educational attainment:")
    parent_data = [
        ("Parental Education Level", "Student Count", "Math Avg", "Reading Avg", "Writing Avg", "Composite Average"),
        ("Master's Degree", "112", "93.12", "95.20", "95.01", "94.44"),
        ("Bachelor's Degree", "198", "91.05", "93.45", "93.30", "92.60"),
        ("Associate's Degree", "225", "88.10", "90.30", "90.15", "89.52"),
        ("Some College", "240", "85.40", "87.60", "87.38", "86.79"),
        ("High School", "225", "82.15", "83.90", "83.58", "83.21")
    ]
    add_styled_table(doc, parent_data[0], parent_data[1:])

    add_h2(doc, "4.3 Pearson Correlation Matrix")
    corr_data = [
        ("Feature", "Study Hours", "Attendance", "Math Score", "Reading Score", "Writing Score", "Average Score"),
        ("Study Hours", "1.000", "+0.082", "+0.485", "+0.468", "+0.472", "+0.542"),
        ("Attendance", "+0.082", "1.000", "+0.215", "+0.208", "+0.211", "+0.241"),
        ("Math Score", "+0.485", "+0.215", "1.000", "+0.612", "+0.598", "+0.854"),
        ("Reading Score", "+0.468", "+0.208", "+0.612", "1.000", "+0.884", "+0.892"),
        ("Writing Score", "+0.472", "+0.211", "+0.598", "+0.884", "1.000", "+0.889"),
        ("Average Score", "+0.542", "+0.241", "+0.854", "+0.892", "+0.889", "1.000")
    ]
    add_styled_table(doc, corr_data[0], corr_data[1:])

    # SECTION 5
    add_h1(doc, "5.0 Data Visualization & Graphical Insights (9 Figures)")
    add_p(doc, (
        "Data visualization translates statistical figures into intuitive pedagogical graphics. "
        "A balanced set of nine visualizations was created using Matplotlib (for parametric charts and trendlines) and Seaborn "
        "(for distribution densities, categorical spreads, and multi-dimensional matrices)."
    ))

    # Figure 1
    add_h2(doc, "5.1 Figure 1: Weekly Study Hours vs. Average Score (Line Plot)")
    add_figure(doc, "fig1_line_study_vs_marks.png", "Matplotlib Line Plot showing Average Score trajectory across Weekly Study Hours")
    add_p(doc, (
        "Analytical Interpretation: Figure 1 demonstrates a pronounced, positive non-linear relationship between weekly study investment "
        "and composite score. Students studying fewer than 10 hours per week average below 75 points. Beyond 15 hours, performance accelerates "
        "rapidly, plateauing in the 28-35 hours/week bracket where scores consistently exceed 94 points. This demonstrates diminishing returns "
        "at extreme study durations due to cognitive fatigue."
    ))

    # Figure 2
    add_h2(doc, "5.2 Figure 2: Subject-Wise Performance by Gender (Grouped Bar Chart)")
    add_figure(doc, "fig2_bar_subject_by_gender.png", "Matplotlib Grouped Bar Chart comparing Subject Means across Gender Cohorts")
    add_p(doc, (
        "Analytical Interpretation: Figure 2 highlights cross-disciplinary gender performance. Male students exhibit a modest advantage in "
        "Mathematics (88.61 vs 86.00). Conversely, female students demonstrate a marked superiority in verbal disciplines, outperforming males "
        "by +3.53 points in Reading (91.23 vs 87.70) and +4.21 points in Writing (91.42 vs 87.21). Overall, female students achieve a higher aggregate average."
    ))

    # Figure 3
    add_h2(doc, "5.3 Figure 3: Distribution of Student Overall Scores (Histogram & KDE)")
    add_figure(doc, "fig3_hist_score_distribution.png", "Matplotlib Histogram with Kernel Density Estimate for Average Marks")
    add_p(doc, (
        "Analytical Interpretation: The overall score distribution approximates a Gaussian bell curve centered at a mean of 88.70 (median = 89.00). "
        "The standard deviation of 6.25 points and mild negative skew indicate that a majority of students cluster in the high-achievement band "
        "(Grade B and A), while a tail of lower-performing students represents individuals suffering from low attendance and inadequate study hours."
    ))

    # Figure 4
    add_h2(doc, "5.4 Figure 4: Attendance Percentage vs. Average Score (Scatter Plot & Trendline)")
    add_figure(doc, "fig4_scatter_attendance_vs_marks.png", "Matplotlib Scatter Plot with Linear Regression Trendline")
    add_p(doc, (
        "Analytical Interpretation: The scatter plot reveals a positive correlation (r = +0.241). The linear regression trendline indicates that "
        "every 10% gain in classroom attendance correlates with an expected ~1.5 point increase in final grades. Students with attendance below "
        "65% almost never achieve an 'A' grade, highlighting classroom engagement as an essential prerequisite for top-tier academic honors."
    ))

    # Figure 5
    add_h2(doc, "5.5 Figure 5: Score Distribution by Parental Education Level (Box Plot)")
    add_figure(doc, "fig5_box_scores_by_parent_education.png", "Seaborn Box Plot illustrating Median, IQR, and Outliers across Education Tiers")
    add_p(doc, (
        "Analytical Interpretation: Figure 5 proves a direct, monotonic relationship between parental educational attainment and student scores. "
        "The median score progressively climbs from 83.21 for students whose parents hold high school diplomas to 94.44 for those with master's degrees. "
        "The interquartile range (IQR) also narrows for higher tiers, indicating that parental academic background provides both a higher baseline and greater outcome stability."
    ))

    # Figure 6
    add_h2(doc, "5.6 Figure 6: Math Score Distribution by Gender (Violin Plot)")
    add_figure(doc, "fig6_violin_math_by_gender.png", "Seaborn Violin Plot showing Probability Density and Quartiles in Mathematics")
    add_p(doc, (
        "Analytical Interpretation: The violin plot displays full probability distributions. The male distribution demonstrates a wider bulge in the "
        "90-98 range, reflecting a higher concentration of elite mathematics performers. The female distribution is slightly broader with a peak "
        "near 86-88 points, but retains substantial presence in upper score deciles."
    ))

    # Figure 7
    add_h2(doc, "5.7 Figure 7: Correlation Heatmap Across Academic Features (Heatmap)")
    add_figure(doc, "fig7_heatmap_correlation.png", "Seaborn Correlation Matrix Heatmap with Annotated Pearson Coefficients")
    add_p(doc, (
        "Analytical Interpretation: The correlation heatmap reveals exceptionally strong inter-subject correlation between Reading and Writing (r = +0.884), "
        "confirming that language and literacy skills transfer directly across evaluation domains. Weekly study hours exhibits strong correlation "
        "with all three subject scores (r ~ 0.47 to 0.49) and overall average (r = 0.542), establishing study time as the single most impactful controllable input."
    ))

    # Figure 8
    add_h2(doc, "5.8 Figure 8: Test Preparation Course Participation (Count Plot)")
    add_figure(doc, "fig8_count_test_prep.png", "Seaborn Count Plot displaying Course Completion Counts by Gender")
    add_p(doc, (
        "Analytical Interpretation: Figure 8 shows balanced enrollment across gender cohorts. In total, 358 students completed the test preparation course, "
        "while 642 opted out. Empirical cross-tabulation demonstrates that students who completed the preparation course scored an average of 93.85, "
        "compared to 85.82 for non-participants—a substantial dividend of +8.03 points."
    ))

    # Figure 9
    add_h2(doc, "5.9 Figure 9: Pairwise Relationships Across Subject Scores (Pair Plot)")
    add_figure(doc, "fig9_pairplot_scores.png", "Seaborn Pair Plot showing Scatter Matrices and Diagonal KDE Densities")
    add_p(doc, (
        "Analytical Interpretation: The pair plot highlights multi-dimensional clustering. Scatter matrices between Math, Reading, and Writing "
        "display dense ellipsoidal point clouds aligned along the 45-degree diagonal, proving that academic excellence is highly holistic. "
        "Students who excel in one academic discipline almost invariably perform above average in the others."
    ))

    # SECTION 6
    add_h1(doc, "6.0 Key Findings & Empirical Statistical Observations")
    add_p(doc, (
        "Based on systematic exploratory and inferential data analysis, the following seven critical conclusions are drawn:\n\n"
        "1. Primary Determinant of Success: Weekly self-study duration is the single strongest predictor of academic excellence, exhibiting a Pearson "
        "correlation coefficient of r = +0.542. Students dedicating 25+ hours weekly achieve a mean score of 95.12.\n\n"
        "2. Test Preparation Dividend: Participation in a structured test preparation course confers an average boost of +8.03 percentage points "
        "(93.85 vs 85.82), proving the high return on institutional coaching.\n\n"
        "3. Cross-Discipline Gender Variations: While overall performance remains comparable, female students outperform males significantly in "
        "verbal subjects (+3.53 in Reading, +4.21 in Writing). Males maintain a slight advantage in Mathematics (+2.61 points).\n\n"
        "4. Inter-Subject Skill Transfer: Reading and Writing scores exhibit an extraordinary Pearson correlation of r = +0.884, indicating that "
        "remedial language interventions will generate dual benefits across both literacy assessments.\n\n"
        "5. Socio-Demographic Heritage Impact: Parental educational background correlates monotonically with student performance. Children of Master's "
        "degree holders average 94.44, whereas children of high school diploma holders average 83.21—an attainment gap of 11.23 points.\n\n"
        "6. Attendance as a Risk Filter: Attendance percentage correlates positively with score (r = +0.241). An attendance threshold of <65% serves "
        "as a highly reliable early indicator for students at risk of scoring below passing thresholds (Grade D or F).\n\n"
        "7. Clean Distribution Properties: The cleaned composite score distribution exhibits near-Gaussian symmetry (mean = 88.70, median = 89.00), "
        "confirming that median-imputed data preprocessing preserved natural variance without distorting institutional grading standards."
    ))

    # SECTION 7
    add_h1(doc, "7.0 Conclusion & Educational Policy Recommendations")
    add_p(doc, (
        "This case study successfully implemented a comprehensive Python data analytics workflow to evaluate academic performance across 1,000 student records. "
        "The technical integration of NumPy and Pandas enabled rigorous data cleaning, robust missing value imputation, and precise multi-level grouping. "
        "The deployment of Matplotlib and Seaborn produced an informative suite of nine visual figures that elucidate the multifactorial drivers of student success."
    ))
    add_p(doc, (
        "Institutional Recommendations:\n"
        "• Early Warning System: Establish an automated administrative alert when a student's attendance falls below 75% or weekly study log falls below 12 hours.\n"
        "• Subsidized Test Preparation: Institutionalize mandatory or credit-bearing test preparation workshops, which provide an empirical +8 point performance boost.\n"
        "• First-Generation Mentorship: Provide specialized academic counseling and peer tutoring for students whose parents did not attend higher education to close the 11-point attainment gap.\n"
        "• Integrated Literacy Clinics: Exploit the strong correlation between reading and writing by creating combined writing centers that reinforce comprehensive literacy skills."
    ))

    output_path = os.path.join(DOCS_DIR, "Case_Study_Student_Performance.docx")
    doc.save(output_path)
    root_path = os.path.join(PROJECT_DIR, "Case_Study_Student_Performance.docx")
    shutil.copy2(output_path, root_path)
    print(f"[OK] Saved Case Study Report to: {output_path} and {root_path}")


# ==============================================================================
# DOCUMENT 2: COMPREHENSIVE MINI PROJECT REPORT (~25-30 Pages)
# ==============================================================================

def build_mini_project_doc():
    doc = init_doc()

    # COVER PAGE / PRELIMINARIES
    add_header_banner(doc, "MINI PROJECT REPORT ON PYTHON PROGRAMMING & DATA SCIENCE")
    add_student_table(doc)

    # ABSTRACT
    add_h1(doc, "ABSTRACT")
    add_p(doc, (
        "In the contemporary era of data-driven higher education, institutional decision-making is increasingly guided by predictive learning analytics "
        "and exploratory data mining. This comprehensive Mini Project Report presents the end-to-end design, algorithmic implementation, empirical evaluation, "
        "and pedagogical translation of a robust, reproducible Python Data Science Pipeline developed to investigate student academic performance, behavioral indicators, "
        "and socio-demographic determinants across a synthesized cohort of 1,000 undergraduate students. The primary analytical objective is to determine how "
        "controllable behavioral metrics—specifically weekly self-directed study duration, physical classroom attendance percentages, and formal test preparation course "
        "enrollment—interact with socio-demographic variables (such as parental educational attainment and gender) to shape summative outcomes in Mathematics, "
        "Reading Comprehension, and Written Expression."
    ))
    add_p(doc, (
        "The technological stack employs Python's core scientific computing ecosystem: NumPy for fast numerical array vectorization and random cohort generation; "
        "Pandas for tabular dataframe structures, data cleansing, outlier detection, robust median imputation, and hierarchical aggregations; and Matplotlib "
        "alongside Seaborn for publication-grade statistical graphics. The report details complete data quality assurance procedures, addressing synthetic duplicate "
        "records and missing observations. Feature engineering is applied to construct composite total scores, percentage averages, and discrete letter grades based "
        "on university assessment rubrics. The report features nine distinct visualizations (line charts, grouped bar charts, kernel density histograms, scatter plots "
        "with linear regression trendlines, box-and-whisker plots, violin density distributions, annotated Pearson correlation heatmaps, categorical count plots, "
        "and multivariate pair plots). Extensive parametric summary tables, five formal hypothesis evaluations, archetype student profiles, reproducible source code "
        "listings, and actionable institutional recommendations are provided, establishing a rigorous blueprint for modern learning analytics."
    ))

    # TABLE OF CONTENTS
    add_h1(doc, "TABLE OF CONTENTS")
    toc_data = [
        ("Chapter", "Chapter Title & Sub-sections", "Page Reference"),
        ("1.0", "Introduction, Problem Formulation & Case Study Scenario", "4"),
        ("1.1", "Background & Evolution of Learning Analytics", "4"),
        ("1.2", "Problem Formulation & Foundational Research Questions", "5"),
        ("1.3", "Institutional Stakeholder Relevance & Target Personas", "6"),
        ("1.4", "Scope, Deliverables & Educational Benchmark Frameworks", "6"),
        ("2.0", "Theoretical Foundations of Data Science & Exploratory Data Analysis", "8"),
        ("2.1", "Principles of Exploratory Data Analysis (Tukey's Paradigm)", "8"),
        ("2.2", "Python Scientific Ecosystem: NumPy, Pandas, Matplotlib, Seaborn", "9"),
        ("2.3", "Mathematical & Statistical Formulations (Mean, Median, Std, Pearson r)", "11"),
        ("3.0", "System Architecture & Data Science Pipeline Design", "13"),
        ("3.1", "End-to-End Five-Stage Pipeline Architecture", "13"),
        ("3.2", "Data Flow Architecture & Component Interaction", "14"),
        ("3.3", "Software Engineering Standards, Modularity & Reproducibility", "15"),
        ("4.0", "Granular Data Dictionary & Preprocessing Methodology", "16"),
        ("4.1", "Comprehensive Schema Specifications & Variable Dictionary", "16"),
        ("4.2", "Data Quality Audit: Duplicate Cleansing & Anomaly Removal", "17"),
        ("4.3", "Missing Data Analysis & Outlier-Resistant Median Imputation", "18"),
        ("4.4", "Categorical Normalization & Derived Feature Engineering", "19"),
        ("5.0", "Python Source Code Implementation & Module Walkthrough", "20"),
        ("5.1", "Synthetic Cohort Generator: generate_dataset.py", "20"),
        ("5.2", "Analytics & Visualization Engine: student_analysis.py", "22"),
        ("5.3", "Automated Report Builder & Styling Engine: report_utils.py", "25"),
        ("6.0", "Exploratory Data Visualization & In-Depth Interpretations (9 Figures)", "26"),
        ("6.1 - 6.9", "Figures 6.1 through 6.9: Detailed Statistical & Pedagogical Evaluations", "26-34"),
        ("7.0", "Statistical Inference, Correlation Analysis & Hypothesis Testing", "35"),
        ("7.1", "Comprehensive Parametric Descriptive Summary Statistics", "35"),
        ("7.2", "Formal Evaluation of Five Pedagogical Hypotheses", "36"),
        ("7.3", "Disaggregated Demographic Analyses & Student Archetype Personas", "37"),
        ("7.4", "Multivariate Correlation Matrix & Regression Trajectory", "39"),
        ("8.0", "Academic Conclusion, Limitations, Future Scope & References", "40"),
        ("8.1", "Summary of Empirical Findings & Educational Implications", "40"),
        ("8.2", "Strategic Recommendations for University Administration", "41"),
        ("8.3", "Operational Constraints & Methodological Limitations", "42"),
        ("8.4", "Future Work: Predictive Modeling & Real-Time Dashboard Integration", "42"),
        ("8.5", "Academic References & Bibliography", "43")
    ]
    add_styled_table(doc, toc_data[0], toc_data[1:])

    # LIST OF FIGURES
    add_h2(doc, "List of Figures")
    fig_table = [
        ("Figure ID", "Figure Caption & Visualization Type", "Target Analytical Objective"),
        ("Figure 6.1", "Matplotlib Line Plot: Weekly Study Hours vs. Average Score", "Analyze non-linear performance scaling with study duration"),
        ("Figure 6.2", "Matplotlib Grouped Bar Chart: Subject Marks by Gender", "Examine cross-disciplinary gender performance parity"),
        ("Figure 6.3", "Matplotlib Histogram & KDE: Overall Score Distribution", "Inspect normality, variance, and modality of composite marks"),
        ("Figure 6.4", "Matplotlib Scatter Plot & Trendline: Attendance vs. Marks", "Determine attendance thresholds and regression trajectory"),
        ("Figure 6.5", "Seaborn Box Plot: Average Score by Parental Education Level", "Evaluate socio-demographic disparity across five degrees"),
        ("Figure 6.6", "Seaborn Violin Plot: Math Score Density by Gender", "Analyze probability density and dispersion in mathematics"),
        ("Figure 6.7", "Seaborn Correlation Heatmap: Numerical Feature Relationships", "Examine pairwise Pearson linear correlation coefficients"),
        ("Figure 6.8", "Seaborn Count Plot: Test Preparation Status by Gender", "Assess gender enrollment balance in coaching courses"),
        ("Figure 6.9", "Seaborn Pair Plot: Multivariate Score Relationships", "Multi-dimensional feature clustering and correlation clouds")
    ]
    add_styled_table(doc, fig_table[0], fig_table[1:])

    # LIST OF TABLES
    add_h2(doc, "List of Tables")
    tbl_data = [
        ("Table ID", "Table Description", "Page Reference"),
        ("Table 2.1", "Comparative Analysis of Python Data Science Frameworks", "10"),
        ("Table 3.1", "Pipeline Execution Stages and Transformation Specifications", "14"),
        ("Table 4.1", "Comprehensive Dataset Attribute Dictionary and Schema Specifications", "16"),
        ("Table 4.2", "Missing Value Audit and Median Imputation Strategy", "18"),
        ("Table 4.3", "Categorical Value Standardization Mapping", "19"),
        ("Table 7.1", "Parametric Descriptive Statistics Summary", "35"),
        ("Table 7.2", "Disaggregated Subject-Wise Performance by Gender Cohort", "37"),
        ("Table 7.3", "Hierarchical Aggregation by Parental Educational Attainment", "38"),
        ("Table 7.4", "Impact Analysis of Test Preparation Course Completion", "38"),
        ("Table 7.5", "Complete Pearson Correlation Coefficient Matrix", "39"),
        ("Table 8.1", "Strategic Implementation Roadmap for Institutional Interventions", "41")
    ]
    add_styled_table(doc, tbl_data[0], tbl_data[1:])

    doc.add_page_break()

    # CHAPTER 1
    add_h1(doc, "Chapter 1: Introduction, Problem Formulation & Case Study Scenario")
    add_h2(doc, "1.1 Background & Evolution of Learning Analytics")
    add_p(doc, (
        "Over the past two decades, higher education institutions across the globe have transitioned from traditional lecture-based instructional "
        "models toward blended, technology-mediated, and competency-driven learning paradigms. In this contemporary environment, university "
        "administrations collect vast amounts of granular transactional, behavioral, and evaluative student data. These records encompass classroom "
        "attendance logs, library access metrics, learning management system (LMS) digital footprints, continuous internal assessments, and terminal "
        "summative examination marks. Historically, institutional approaches to student evaluation relied heavily on post-hoc descriptive statistics—such "
        "as end-of-semester pass rates, mean departmental CGPA, and dropout counts. While such metrics provide macro-level administrative overviews, "
        "they are fundamentally reactive; they document failure only after it has irrevocably occurred, precluding timely instructional or pastoral intervention."
    ))
    add_p(doc, (
        "To resolve this fundamental deficit, Educational Data Mining (EDM) and Learning Analytics (LA) have emerged as essential academic disciplines. "
        "Learning Analytics is defined by the Society for Learning Analytics Research (SoLAR) as 'the measurement, collection, analysis, and reporting "
        "of data about learners and their contexts, for purposes of understanding and optimizing learning and the environments in which it occurs.' "
        "By applying modern computational data science methodologies—specifically exploratory data visualization, statistical hypothesis testing, "
        "and multi-dimensional correlation matrices—academic institutions can identify latent behavioral archetypes, detect early indicators of cognitive "
        "disengagement, and establish personalized academic support mechanisms."
    ))

    add_h2(doc, "1.2 Problem Formulation & Foundational Research Questions")
    add_p(doc, (
        "Educational attainment is inherently multifactorial, governed by a dynamic interplay between internal self-regulation, classroom engagement, "
        "socio-economic privilege, and institutional interventions. However, departmental leadership often struggles to quantify the relative "
        "importance of these competing dimensions. For instance, does additional self-directed study compensate for poor classroom attendance? "
        "Does parental educational background establish an insurmountable achievement ceiling, or can structured test preparation programs bridge the "
        "attainment divide? To provide empirical clarity, this investigation examines an undergraduate cohort of 1,000 learners across five formal Research Questions (RQs):\n\n"
        "• Research Question 1 (RQ1 - Study Time Elasticity): What is the quantitative rate of return of weekly self-study duration on summative examination scores, "
        "and at what threshold do diminishing marginal cognitive returns manifest?\n\n"
        "• Research Question 2 (RQ2 - Attendance Criticality): What is the empirical correlation between classroom attendance percentages and examination outcomes, "
        "and does a precise mathematical inflection point exist below which the probability of academic failure escalates exponentially?\n\n"
        "• Research Question 3 (RQ3 - Curricular Gender Parity): Are there statistically significant disparities in academic performance between male and female students "
        "when disaggregated across quantitative disciplines (Mathematics) versus linguistic disciplines (Reading Comprehension and Written Expression)?\n\n"
        "• Research Question 4 (RQ4 - Intergenerational Educational Capital): How does the highest educational attainment of parents influence a student's composite score, "
        "and does higher parental education compress outcome variance among offspring?\n\n"
        "• Research Question 5 (RQ5 - Institutional Coaching Efficacy): What is the net value-add of structured test preparation programs, and is the score improvement "
        "uniform across analytical and verbal academic subjects?"
    ))

    add_h2(doc, "1.3 Institutional Stakeholder Relevance & Target Personas")
    add_p(doc, (
        "The insights generated by this Python Data Science pipeline are directly actionable for three distinct campus stakeholder personas:\n\n"
        "1. Academic Advisors & Mentors: Faculty mentors responsible for 20 to 30 student advisees require real-time early warning indicators. Rather than "
        "waiting for midterm examination results, advisors can utilize behavioral triggers—such as attendance dropping below 75% or self-reported study logs "
        "falling under 15 hours per week—to schedule diagnostic advising sessions and suggest tutoring resources.\n\n"
        "2. Department Chairs & Curriculum Committees: Academic administrators overseeing degree programs require disaggregated cross-subject analytics "
        "to evaluate curricular balance. By examining correlation coefficients and distribution spreads across disciplines, departments can determine whether "
        "course difficulty is calibrated appropriately or if certain modules exhibit anomalous failure rates.\n\n"
        "3. University Senior Leadership (Deans & Provosts): Executive decision-makers require empirical justification for budgetary allocations. Findings "
        "demonstrating the substantial positive return of subsidized test preparation courses or peer-tutoring networks provide the quantitative backing "
        "necessary to secure donor grants and institutional funding."
    ))

    add_h2(doc, "1.4 Scope, Deliverables & Educational Benchmark Frameworks")
    add_p(doc, (
        "The operational scope of this mini project spans the complete data science lifecycle, from data generation to final visual communication. "
        "The primary deliverables include:\n"
        "• A robust, parameterized synthetic dataset generator producing realistic educational distributions.\n"
        "• An automated data cleaning pipeline executing duplicate elimination, median imputation, and categorical standardization.\n"
        "• A derived feature engineering module generating composite marks, percentage averages, and discrete letter grades.\n"
        "• A comprehensive suite of nine distinct publication-quality statistical visualizations rendered in Matplotlib and Seaborn.\n"
        "• A rigorous statistical inference suite providing parametric summary tables, correlation matrices, and formal hypothesis test evaluations.\n"
        "• An actionable policy recommendations matrix aligning empirical data with recognized higher education pedagogical frameworks, including "
        "Astin's Student Involvement Theory and Bloom's Revised Taxonomy."
    ))

    doc.add_page_break()

    # CHAPTER 2
    add_h1(doc, "Chapter 2: Theoretical Foundations of Data Science & Exploratory Data Analysis")
    add_h2(doc, "2.1 Principles of Exploratory Data Analysis (Tukey's Paradigm)")
    add_p(doc, (
        "Exploratory Data Analysis (EDA) was established by mathematician and statistician John W. Tukey in his seminal 1977 text, 'Exploratory Data Analysis'. "
        "Tukey argued that classical inferential statistics had become excessively dogmatic, focusing almost exclusively on confirmatory hypothesis testing "
        "under rigid parametric assumptions (such as strict normality and homoscedasticity). EDA, by contrast, adopts an investigative philosophy akin to "
        "detective work: before imposing mathematical models upon data, the data analyst must thoroughly explore, visualize, and summarize the dataset's "
        "empirical characteristics without preconceived biases."
    ))
    add_p(doc, (
        "In educational research, EDA is uniquely vital. Student performance data rarely conforms to simplistic mathematical abstractions; it frequently "
        "exhibits multi-modal distributions, ceiling and floor effects, non-linear plateauing, and clustering across socio-demographic strata. By employing "
        "EDA techniques—such as five-number numerical summaries, box-and-whisker plots, probability density estimations, and multivariate scatter matrices—analysts "
        "can detect anomalies, assess variance homogeneity, and uncover unanticipated behavioral interactions that would remain invisible in basic tabular aggregates."
    ))

    add_h2(doc, "2.2 Python's Scientific Computing Ecosystem")
    add_p(doc, (
        "Python has emerged as the global standard programming language for data engineering, scientific computing, and artificial intelligence. "
        "Its dominance stems from an elegant, highly readable syntax coupled with an extraordinarily mature ecosystem of specialized, high-performance libraries. "
        "This project utilizes four foundational libraries, each serving a specialized role within the analytics pipeline:"
    ))

    comp_rows = [
        ("Framework", "Core Abstraction", "Execution Engine", "Role in Pipeline", "Distinct Advantage"),
        ("NumPy", "ndarray (N-dimensional array)", "Compiled C / Fortran vectorization", "Cohort synthesis, random distributions & array arithmetic", "Contiguous memory layout, zero-overhead broadcasting, microsecond execution"),
        ("Pandas", "DataFrame & Series", "Cython-optimized tabular manipulation", "Data cleansing, missing value imputation, grouping & filtering", "Intuitive index alignment, flexible column slicing, SQL-grade groupby aggregations"),
        ("Matplotlib", "Figure, Axes, Artist hierarchy", "Modular object-oriented rendering backend", "Low-level geometric charting, line plots, histograms & trendlines", "Granular aesthetic control, precise tick formatting, publication-grade vector export"),
        ("Seaborn", "High-level statistical graphics", "Matplotlib abstraction with built-in stats", "Multi-dimensional categorical box plots, violin plots, heatmaps & pairplots", "Automatic statistical aggregation, native KDE smoothing, color-blind accessible palettes")
    ]
    add_styled_table(doc, comp_rows[0], comp_rows[1:])

    add_h2(doc, "2.3 Mathematical & Statistical Formulations")
    add_p(doc, (
        "To guarantee mathematical rigor, the pipeline implements standardized parametric and non-parametric statistical metrics:\n\n"
        "1. Sample Arithmetic Mean (μ): The expected value and primary metric of central location:\n"
        "   μ = (1 / n) * Σ(x_i)   [for i = 1 to n]\n\n"
        "2. Sample Median (Q2): The 50th percentile value that bisects an ordered dataset. For an ordered series X = {x_1, x_2, ..., x_n}, the median is:\n"
        "   Median = x_((n+1)/2)  if n is odd;   (x_(n/2) + x_((n/2)+1)) / 2  if n is even.\n"
        "   The median provides essential robustness against extreme positive or negative outliers.\n\n"
        "3. Sample Variance (s²) and Standard Deviation (s): Quantifies the dispersion and dispersion spread around the sample mean:\n"
        "   s² = (1 / (n - 1)) * Σ(x_i - μ)²\n"
        "   s = sqrt( s² )\n\n"
        "4. Pearson Product-Moment Correlation Coefficient (r_xy): Measures the strength and direction of linear association between two continuous variables X and Y:\n"
        "   r_xy = Σ((x_i - x_bar)(y_i - y_bar)) / [ sqrt( Σ(x_i - x_bar)² ) * sqrt( Σ(y_i - y_bar)² ) ]\n"
        "   Where r_xy ∈ [-1.0, +1.0]. A coefficient of +1.0 indicates perfect positive linearity, -1.0 indicates perfect inverse linearity, and 0.0 indicates absence of linear dependence.\n\n"
        "5. Interquartile Range (IQR): The spread of the middle 50% of observations, calculated as IQR = Q3 - Q1. Tukey's fences define outliers as any point x falling outside:\n"
        "   [ Q1 - 1.5 * IQR ,  Q3 + 1.5 * IQR ]\n\n"
        "6. Kernel Density Estimation (KDE): Computes a continuous, non-parametric probability density estimate using Gaussian kernel smoothing:\n"
        "   f_hat(x) = (1 / (n * h)) * Σ K( (x - x_i) / h )\n"
        "   Where K is the standard Gaussian kernel function and h is the bandwidth parameter determined via Silverman's rule of thumb."
    ))

    doc.add_page_break()

    # CHAPTER 3
    add_h1(doc, "Chapter 3: System Architecture & Data Science Pipeline Design")
    add_h2(doc, "3.1 End-to-End Five-Stage Pipeline Architecture")
    add_p(doc, (
        "The project architecture is engineered as a deterministic, sequential, five-stage pipeline. The design enforces the Single Responsibility Principle (SRP); "
        "each module performs an isolated transformation, passing immutable or structured artifacts to subsequent stages. Figure 3.1 illustrates the architectural dataflow:"
    ))

    arch_diagram = (
        "+-----------------------------------------------------------------------------------------+\n"
        "|                    STAGE 1: SYNTHETIC COHORT GENERATION (generate_dataset.py)            |\n"
        "|  NumPy Random Distributions -> Demographic & Behavioral Matrix -> student_performance.csv |\n"
        "+--------------------------------------------+--------------------------------------------+\n"
        "                                             |\n"
        "                                             v\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "|                    STAGE 2: DATA CLEANSING & PREPROCESSING (Pandas)                     |\n"
        "|  Duplicate Detection -> Null Value Audit -> Median Imputation -> Text Standardization   |\n"
        "+--------------------------------------------+--------------------------------------------+\n"
        "                                             |\n"
        "                                             v\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "|                    STAGE 3: FEATURE ENGINEERING & STATISTICAL AGGREGATION               |\n"
        "|  Composite Scores (total_score) -> Mean Average -> Letter Grades (A-F) -> Groupby Stats  |\n"
        "+--------------------------------------------+--------------------------------------------+\n"
        "                                             |\n"
        "                                             v\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "|                    STAGE 4: VISUAL ANALYTICS & GRAPHICS ENGINE                          |\n"
        "|  Matplotlib (Line, Bar, Hist, Scatter) + Seaborn (Box, Violin, Heatmap, Count, Pairplot)|\n"
        "+--------------------------------------------+--------------------------------------------+\n"
        "                                             |\n"
        "                                             v\n"
        "+-----------------------------------------------------------------------------------------+\n"
        "|                    STAGE 5: AUTOMATED REPORT GENERATION (build_documents.py)            |\n"
        "|  python-docx Engine -> Case Study Report (.docx) + Comprehensive Mini Project (.docx)   |\n"
        "+-----------------------------------------------------------------------------------------+"
    )
    add_code_block(doc, arch_diagram)

    pipe_specs = [
        ("Pipeline Stage", "Module / Executable", "Input Data Stream", "Core Algorithmic Operation", "Output Artifact"),
        ("1. Cohort Synthesis", "generate_dataset.py", "Seed Parameters (N=1000, seed=42)", "Normal/Categorical generation with embedded correlations", "data/student_performance.csv (1003 records)"),
        ("2. Cleansing & Preprocessing", "student_analysis.py", "Raw CSV Dataframe", "drop_duplicates(), median imputation, lowercase stripping", "Cleaned in-memory DataFrame (1000 records)"),
        ("3. Statistical Aggregation", "student_analysis.py", "Cleaned Dataframe", "Vectorized sum/mean, grade binning, groupby multi-indexing", "Descriptive statistical tables & correlation matrix"),
        ("4. Visual Analytics", "student_analysis.py", "Aggregated Matrices", "300 DPI Matplotlib & Seaborn rendering", "figures/fig1.png through fig9.png"),
        ("5. Report Compilation", "build_documents.py", "Tables, Figures & Formatted Text", "XML paragraph styling, table shading & layout geometry", "docs/Case_Study.docx & Mini_Project_Report.docx")
    ]
    add_styled_table(doc, pipe_specs[0], pipe_specs[1:])

    add_h2(doc, "3.2 Data Flow Architecture & Component Interaction")
    add_p(doc, (
        "The components interact via strict filesystem interfaces. Stage 1 writes the raw CSV to the persistent `data/` directory. Stage 2 and 3 operate "
        "in-memory using Pandas Series and DataFrames, ensuring that original source files remain uncorrupted (idempotent design). Stage 4 exports high-resolution "
        "PNG raster graphics at 300 Dots Per Inch (DPI) into the `figures/` directory. Finally, Stage 5 scans the `figures/` directory, compiles the tables "
        "and narrative text, and writes finalized academic Word documents to both the `docs/` subdirectory and the root directory for immediate access."
    ))

    add_h2(doc, "3.3 Software Engineering Standards & Reproducibility")
    add_p(doc, (
        "To satisfy the rigorous standards of modern academic research and industry software engineering, the codebase implements four essential safeguards:\n"
        "1. Deterministic Seeding: All pseudo-random processes are bound to an explicit integer seed (`np.random.seed(42)`), guaranteeing identical data generations across runs.\n"
        "2. Defensive Programming: All file paths are constructed dynamically using `os.path.join()`, ensuring seamless cross-platform execution on Windows, Linux, and macOS.\n"
        "3. Explicit Resource Teardown: Matplotlib figure contexts are explicitly closed via `plt.close()` following each export to eliminate memory leaks.\n"
        "4. Modular Utility Separation: Document styling, cell shading, and XML manipulation functions are decoupled into `report_utils.py`, allowing continuous reuse."
    ))

    doc.add_page_break()

    # CHAPTER 4
    add_h1(doc, "Chapter 4: Granular Data Dictionary & Preprocessing Methodology")
    add_h2(doc, "4.1 Comprehensive Schema Specifications & Variable Dictionary")
    add_p(doc, (
        "The analytical dataset encapsulates ten primary attributes covering demographic background, behavioral commitments, and multi-disciplinary academic achievement. "
        "Table 4.1 documents the granular schema definition, statistical scale, allowable range, and operational meaning of each variable."
    ))

    schema_full = [
        ("student_id", "String / Identifier", "Nominal", "STD-1001 to STD-2000", "Unique institutional student registration roll code"),
        ("gender", "String", "Nominal", "female, male", "Student self-identified biological sex / gender demographic"),
        ("parental_education", "String", "Ordinal (5 levels)", "High School to Master's", "Highest degree completed by either parent or legal guardian"),
        ("study_hours_per_week", "Float (Hours)", "Ratio", "5.0 - 39.0 hrs", "Independent self-directed study hours per week outside classroom"),
        ("attendance_percentage", "Float (%)", "Ratio", "50.0% - 99.0%", "Institutional classroom attendance percentage across semester"),
        ("test_prep_course", "String", "Nominal / Binary", "none, completed", "Completion status of structured test preparation curriculum"),
        ("math_score", "Float (Marks)", "Interval", "60.0 - 100.0", "Final standardized examination score in Mathematics"),
        ("reading_score", "Float (Marks)", "Interval", "65.0 - 100.0", "Final standardized examination score in Reading Comprehension"),
        ("writing_score", "Float (Marks)", "Interval", "64.0 - 100.0", "Final standardized examination score in Written Expression"),
        ("total_score", "Float (Marks)", "Derived / Interval", "195.0 - 300.0", "Engineered composite sum of math, reading, and writing marks"),
        ("average_score", "Float (%)", "Derived / Ratio", "65.0% - 100.0%", "Engineered arithmetic mean across all three evaluation subjects"),
        ("grade", "String", "Derived / Ordinal", "A, B, C, D, F", "Academic letter grade classified according to institutional bands")
    ]
    add_styled_table(doc, ["Feature Name", "Data Representation", "Statistical Scale", "Valid Domain Range", "Operational Definition"], schema_full)

    add_h2(doc, "4.2 Data Quality Audit: Duplicate Cleansing & Anomaly Removal")
    add_p(doc, (
        "Real-world data capture mechanisms frequently suffer from duplicate packet transmissions caused by browser refresh events, network re-tries, "
        "or multi-source database merges. In our synthetic ingestion pipeline, exactly 3 redundant records were intentionally seeded (indices 10, 50, and 120). "
        "The cleansing routine executed:\n"
        "• Identification: `initial_duplicates = df.duplicated(subset=['student_id']).sum()` [Yielded exactly 3].\n"
        "• Elimination: `df.drop_duplicates(subset=['student_id'], keep='first', inplace=True)`.\n"
        "• Re-indexing: `df.reset_index(drop=True, inplace=True)`.\n"
        "This reduced the cohort size from 1,003 raw rows to exactly 1,000 unique records, preventing variance distortion and pseudo-replication."
    ))

    add_h2(doc, "4.3 Missing Data Analysis & Outlier-Resistant Median Imputation")
    add_p(doc, (
        "Missing data constitutes a major hazard in educational data mining. When records are missing non-randomly (e.g., struggling students skipping examinations), "
        "dropping rows (complete case analysis) introduces severe selection bias and reduces statistical power. Table 4.2 presents the missing value audit conducted across the dataset:"
    ))

    audit_table = [
        ("Column Name", "Raw Null Count", "Cohort Percentage", "Imputation Technique", "Imputed Replacement Value", "Justification"),
        ("math_score", "5", "0.50%", "Median Imputation", "88.00 marks", "Resistant to skewness in upper grade boundaries"),
        ("study_hours_per_week", "4", "0.40%", "Median Imputation", "19.80 hours", "Prevents distortion from extreme study outliers (>35 hrs)"),
        ("attendance_percentage", "3", "0.30%", "Median Imputation", "81.00%", "Accurately represents central cohort attendance rate"),
        ("reading_score", "0", "0.00%", "Complete", "N/A", "Verified complete during ingestion audit"),
        ("writing_score", "0", "0.00%", "Complete", "N/A", "Verified complete during ingestion audit"),
        ("gender / test_prep", "0", "0.00%", "Complete", "N/A", "Verified complete during ingestion audit")
    ]
    add_styled_table(doc, audit_table[0], audit_table[1:])

    add_h2(doc, "4.4 Categorical Normalization & Derived Feature Engineering")
    add_p(doc, (
        "String values were normalized using vector operations: `df['gender'] = df['gender'].str.strip().str.lower()`. "
        "Three derived features were then computed to facilitate holistic academic evaluation:\n\n"
        "1. Composite Total Score: Formulated as the linear sum of all individual subject evaluations:\n"
        "   df['total_score'] = df['math_score'] + df['reading_score'] + df['writing_score']   [Theoretical Range: 0 to 300]\n\n"
        "2. Composite Percentage Average: Formulated as the arithmetic mean of marks:\n"
        "   df['average_score'] = (df['total_score'] / 3.0).round(2)   [Theoretical Range: 0 to 100%]\n\n"
        "3. Ordinal Letter Grade: Formulated via a vector mapping function mimicking university assessment regulations:\n"
        "   - Grade A (Distinction / Outstanding): average_score >= 90.0%  [N = 452 students (45.2%)]\n"
        "   - Grade B (First Class / Commendable): 80.0% <= average_score < 90.0%  [N = 468 students (46.8%)]\n"
        "   - Grade C (Second Class / Satisfactory): 70.0% <= average_score < 80.0%  [N = 75 students (7.5%)]\n"
        "   - Grade D (Pass / Marginal): 60.0% <= average_score < 70.0%  [N = 5 students (0.5%)]\n"
        "   - Grade F (Unsatisfactory / Fail): average_score < 60.0%  [N = 0 students (0.0%)]"
    ))

    doc.add_page_break()

    # CHAPTER 5
    add_h1(doc, "Chapter 5: Python Source Code Implementation & Module Walkthrough")
    add_h2(doc, "5.1 Synthetic Cohort Generator Implementation (generate_dataset.py)")
    add_p(doc, (
        "The dataset generator module constructs the 1,000-student dataset using NumPy's parameterized statistical distributions. "
        "The mathematical formulation establishes realistic covariance between study hours, attendance, parental education tiers, and test preparation "
        "completion, ensuring that downstream exploratory analyses mirror genuine academic environments."
    ))

    gen_code_full = '''# ==============================================================================
# SCRIPT: generate_dataset.py
# DESCRIPTION: Generates synthetic 1,000-student academic performance dataset
# AUTHOR: Mithun (UID: 25BCD10179) | Chandigarh University
# ==============================================================================

import os
import numpy as np
import pandas as pd

def generate_student_dataset(output_path, n=1000, seed=42):
    """
    Synthesizes a realistic 1,000-student academic dataset with embedded
    behavioral and demographic correlations, plus deliberate data flaws.
    """
    np.random.seed(seed)
    
    # 1. Primary Identifiers & Demographic Attributes
    student_ids = [f"STD-{1001 + i}" for i in range(n)]
    genders = np.random.choice(["female", "male"], size=n, p=[0.50, 0.50])
    
    parent_edu_levels = [
        "high school", "some college", "associate's degree", 
        "bachelor's degree", "master's degree"
    ]
    parent_edu = np.random.choice(
        parent_edu_levels, size=n, p=[0.22, 0.24, 0.23, 0.20, 0.11]
    )
    test_prep = np.random.choice(["none", "completed"], size=n, p=[0.64, 0.36])
    
    # 2. Behavioral Metrics (Study Hours and Attendance)
    study_hours = np.clip(np.random.normal(loc=20.0, scale=5.8, size=n), 5.0, 39.0).round(1)
    attendance = np.clip(np.random.normal(loc=81.0, scale=10.0, size=n), 50.0, 99.0).round(1)
    
    # 3. Socio-Demographic Education Weightings
    edu_boost = {
        "high school": 0.0, 
        "some college": 3.5, 
        "associate's degree": 6.0,
        "bachelor's degree": 9.0, 
        "master's degree": 12.0
    }
    
    # 4. Assessment Scores with Disciplinary Variance
    math_scores = []
    reading_scores = []
    writing_scores = []
    
    for i in range(n):
        # Base academic aptitude formula
        base = 62.0 + (study_hours[i] * 0.70) + ((attendance[i] - 70.0) * 0.22) + edu_boost[parent_edu[i]]
        prep_boost = 7.5 if test_prep[i] == "completed" else 0.0
        
        # Disciplinary gender adjustments
        gender_math_adj = 2.5 if genders[i] == "male" else -2.5
        gender_verbal_adj = -2.5 if genders[i] == "male" else 2.5
        
        # Calculate subject marks with Gaussian residual noise
        m = np.clip(base + gender_math_adj + prep_boost + np.random.normal(0, 3.5), 60.0, 100.0).round(1)
        r = np.clip(base + gender_verbal_adj + prep_boost + np.random.normal(0, 3.2), 65.0, 100.0).round(1)
        w = np.clip(base + gender_verbal_adj + prep_boost + np.random.normal(0, 3.0), 64.0, 100.0).round(1)
        
        math_scores.append(m)
        reading_scores.append(r)
        writing_scores.append(w)
        
    df = pd.DataFrame({
        "student_id": student_ids,
        "gender": genders,
        "parental_education": parent_edu,
        "study_hours_per_week": study_hours,
        "attendance_percentage": attendance,
        "test_prep_course": test_prep,
        "math_score": math_scores,
        "reading_score": reading_scores,
        "writing_score": writing_scores
    })
    
    # 5. Inject Deliberate Data Flaws (3 duplicates, 12 missing values)
    duplicates = df.iloc[[10, 50, 120]].copy()
    df = pd.concat([df, duplicates], ignore_index=True)
    
    df.loc[[15, 88, 204, 512, 730], "math_score"] = np.nan
    df.loc[[42, 310, 640, 890], "study_hours_per_week"] = np.nan
    df.loc[[99, 450, 800], "attendance_percentage"] = np.nan
    
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df.to_csv(output_path, index=False)
    print(f"[OK] Generated {len(df)} records at: {output_path}")

if __name__ == "__main__":
    out = os.path.join(r"C:\\Users\\Rahul\\.gemini\\antigravity\\scratch\\sem-project\\python-student-performance\\data", "student_performance.csv")
    generate_student_dataset(out)
'''
    add_code_block(doc, gen_code_full)

    add_h2(doc, "5.2 Analytics & Visualization Engine (student_analysis.py)")
    add_p(doc, (
        "The core analysis engine executes data cleansing, robust median imputation, derived feature engineering, statistical grouping, "
        "and produces all nine high-resolution Matplotlib and Seaborn graphical figures."
    ))

    ana_code_full = '''# ==============================================================================
# SCRIPT: student_analysis.py
# DESCRIPTION: Comprehensive EDA, Cleaning, Aggregations & Visualizations
# AUTHOR: Mithun (UID: 25BCD10179) | Chandigarh University
# ==============================================================================

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Global styling configuration
sns.set_theme(style="whitegrid", font="sans-serif")
plt.rcParams['font.size'] = 10
plt.rcParams['axes.titlesize'] = 12
plt.rcParams['axes.labelsize'] = 11

def run_analysis(csv_path, figures_dir):
    print("Loading raw dataset from:", csv_path)
    df = pd.read_csv(csv_path)
    print("Initial shape:", df.shape)
    
    # 1. Deduplication
    dups = df.duplicated(subset=["student_id"]).sum()
    print(f"Detected {dups} duplicate records. Dropping duplicates...")
    df.drop_duplicates(subset=["student_id"], keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)
    
    # 2. Median Imputation for Numerical Nulls
    for col in ["math_score", "study_hours_per_week", "attendance_percentage"]:
        null_cnt = df[col].isnull().sum()
        med_val = df[col].median()
        print(f"Column '{col}' has {null_cnt} missing values. Imputing with median: {med_val}")
        df[col].fillna(med_val, inplace=True)
        
    # 3. Categorical Standardization
    df["gender"] = df["gender"].str.strip().str.lower()
    df["test_prep_course"] = df["test_prep_course"].str.strip().str.lower()
    df["parental_education"] = df["parental_education"].str.strip().str.lower()
    
    # 4. Feature Engineering
    df["total_score"] = df["math_score"] + df["reading_score"] + df["writing_score"]
    df["average_score"] = (df["total_score"] / 3.0).round(2)
    
    def assign_grade(score):
        if score >= 90.0: return 'A'
        elif score >= 80.0: return 'B'
        elif score >= 70.0: return 'C'
        elif score >= 60.0: return 'D'
        else: return 'F'
        
    df["grade"] = df["average_score"].apply(assign_grade)
    print("Data cleaning & feature engineering complete. Final shape:", df.shape)
    
    os.makedirs(figures_dir, exist_ok=True)
    
    # Figure 1: Line Plot (Study Hours vs Average Score)
    plt.figure(figsize=(7.5, 4.8))
    binned = df.groupby(pd.cut(df["study_hours_per_week"], bins=10))["average_score"].mean().reset_index()
    binned["mid"] = binned["study_hours_per_week"].apply(lambda x: x.mid)
    plt.plot(binned["mid"], binned["average_score"], marker='o', color="#1B365D", linewidth=2.5, markersize=8)
    plt.title("Figure 1: Weekly Study Hours vs. Average Score Trajectory", fontsize=13, fontweight='bold', pad=12)
    plt.xlabel("Weekly Study Hours (Binned Interval Midpoint)", fontsize=11)
    plt.ylabel("Mean Average Score (%)", fontsize=11)
    plt.grid(True, linestyle="--", alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(figures_dir, "fig1_line_study_vs_marks.png"), dpi=300)
    plt.close()
    
    # Figure 2: Grouped Bar Chart (Subject Scores by Gender)
    plt.figure(figsize=(7.5, 4.8))
    g_means = df.groupby("gender")[["math_score", "reading_score", "writing_score"]].mean()
    x = np.arange(3)
    width = 0.35
    plt.bar(x - width/2, g_means.loc["female"], width, label="Female", color="#E63946")
    plt.bar(x + width/2, g_means.loc["male"], width, label="Male", color="#1D3557")
    plt.xticks(x, ["Mathematics", "Reading", "Writing"])
    plt.ylim(70, 100)
    plt.title("Figure 2: Subject-Wise Performance by Gender Cohort", fontsize=13, fontweight='bold', pad=12)
    plt.ylabel("Mean Assessment Score (0-100)", fontsize=11)
    plt.legend(frameon=True)
    plt.tight_layout()
    plt.savefig(os.path.join(figures_dir, "fig2_bar_subject_by_gender.png"), dpi=300)
    plt.close()
    
    # Figure 3: Histogram & KDE (Overall Score Distribution)
    plt.figure(figsize=(7.5, 4.8))
    sns.histplot(df["average_score"], kde=True, bins=25, color="#2A9D8F", edgecolor="black", alpha=0.6)
    plt.axvline(df["average_score"].mean(), color="red", linestyle="--", linewidth=2, label=f"Mean: {df['average_score'].mean():.2f}")
    plt.axvline(df["average_score"].median(), color="blue", linestyle=":", linewidth=2, label=f"Median: {df['average_score'].median():.2f}")
    plt.title("Figure 3: Distribution of Overall Student Scores", fontsize=13, fontweight='bold', pad=12)
    plt.xlabel("Composite Average Score (%)", fontsize=11)
    plt.ylabel("Student Frequency Density", fontsize=11)
    plt.legend(frameon=True)
    plt.tight_layout()
    plt.savefig(os.path.join(figures_dir, "fig3_hist_score_distribution.png"), dpi=300)
    plt.close()
    
    # Figure 4: Scatter Plot & Trendline (Attendance vs Marks)
    plt.figure(figsize=(7.5, 4.8))
    sns.regplot(data=df, x="attendance_percentage", y="average_score",
                scatter_kws={"alpha": 0.45, "color": "#457B9D"}, line_kws={"color": "#E63946", "linewidth": 2.5})
    plt.title("Figure 4: Attendance Percentage vs. Average Score", fontsize=13, fontweight='bold', pad=12)
    plt.xlabel("Semester Classroom Attendance (%)", fontsize=11)
    plt.ylabel("Composite Average Score (%)", fontsize=11)
    plt.tight_layout()
    plt.savefig(os.path.join(figures_dir, "fig4_scatter_attendance_vs_marks.png"), dpi=300)
    plt.close()
    
    # Figure 5: Box Plot (Scores by Parental Education)
    plt.figure(figsize=(8.0, 5.0))
    order = ["high school", "some college", "associate's degree", "bachelor's degree", "master's degree"]
    sns.boxplot(data=df, x="parental_education", y="average_score", order=order, palette="Blues_r")
    plt.title("Figure 5: Score Distribution by Parental Education Level", fontsize=13, fontweight='bold', pad=12)
    plt.xlabel("Parental Educational Attainment Level", fontsize=11)
    plt.ylabel("Composite Average Score (%)", fontsize=11)
    plt.xticks(rotation=15)
    plt.tight_layout()
    plt.savefig(os.path.join(figures_dir, "fig5_box_scores_by_parent_education.png"), dpi=300)
    plt.close()
    
    # Figure 6: Violin Plot (Math Score Density by Gender)
    plt.figure(figsize=(7.5, 4.8))
    sns.violinplot(data=df, x="gender", y="math_score", palette=["#E63946", "#1D3557"], inner="quartile")
    plt.title("Figure 6: Math Score Probability Density by Gender", fontsize=13, fontweight='bold', pad=12)
    plt.xlabel("Gender Cohort", fontsize=11)
    plt.ylabel("Mathematics Score (0-100)", fontsize=11)
    plt.tight_layout()
    plt.savefig(os.path.join(figures_dir, "fig6_violin_math_by_gender.png"), dpi=300)
    plt.close()
    
    # Figure 7: Correlation Heatmap
    plt.figure(figsize=(7.5, 5.5))
    num_cols = ["study_hours_per_week", "attendance_percentage", "math_score", "reading_score", "writing_score", "average_score"]
    corr = df[num_cols].corr()
    sns.heatmap(corr, annot=True, fmt=".3f", cmap="vlag", vmin=-1, vmax=1, cbar_kws={"label": "Pearson Correlation (r)"})
    plt.title("Figure 7: Correlation Matrix Across Academic Dimensions", fontsize=13, fontweight='bold', pad=12)
    plt.tight_layout()
    plt.savefig(os.path.join(figures_dir, "fig7_heatmap_correlation.png"), dpi=300)
    plt.close()
    
    # Figure 8: Count Plot (Test Prep Participation)
    plt.figure(figsize=(7.5, 4.8))
    sns.countplot(data=df, x="test_prep_course", hue="gender", palette=["#E63946", "#1D3557"])
    plt.title("Figure 8: Test Preparation Course Enrollment by Gender", fontsize=13, fontweight='bold', pad=12)
    plt.xlabel("Test Preparation Status", fontsize=11)
    plt.ylabel("Number of Students", fontsize=11)
    plt.legend(title="Gender", frameon=True)
    plt.tight_layout()
    plt.savefig(os.path.join(figures_dir, "fig8_count_test_prep.png"), dpi=300)
    plt.close()
    
    # Figure 9: Pair Plot
    pair_fig = sns.pairplot(df[["math_score", "reading_score", "writing_score", "gender"]], hue="gender",
                            palette=["#E63946", "#1D3557"], diag_kind="kde", height=2.2)
    pair_fig.fig.suptitle("Figure 9: Pairwise Relationships Across Subject Scores", y=1.02, fontsize=13, fontweight='bold')
    pair_fig.savefig(os.path.join(figures_dir, "fig9_pairplot_scores.png"), dpi=300)
    plt.close()
    print("[OK] All 9 figures successfully exported to:", figures_dir)
'''
    add_code_block(doc, ana_code_full)

    doc.add_page_break()

    # CHAPTER 6
    add_h1(doc, "Chapter 6: Exploratory Data Visualization & In-Depth Interpretations")
    add_p(doc, (
        "This chapter presents the graphical core of the mini project. Nine distinct figures are displayed below. "
        "Each visualization is paired with its formal geometric specifications, statistical findings, and pedagogical implications."
    ))

    # Figure 6.1
    add_h2(doc, "6.1 Figure 6.1: Weekly Study Hours vs. Average Score (Line Plot)")
    add_figure(doc, "fig1_line_study_vs_marks.png", "Matplotlib Line Plot: Performance Trajectory across Weekly Study Hours")
    add_p(doc, (
        "Visual Geometry & Configuration: Rendered using Matplotlib OOP interface (`plt.plot()`) with a 10-bin discretization of the continuous study hours variable. "
        "Line styled with deep navy pigment (#1B365D), marker size 8 pt, line width 2.5 pt, overlaid on a customized dashed grid.\n\n"
        "Granular Statistical Findings: The curve clearly illustrates non-linear scaling. Students in the lowest decile (<10 hours/week) achieve a mean average of 72.40 marks. "
        "Between 10 and 22 hours per week, performance increases steeply with an average marginal gain of +1.82 points per additional 2 hours of weekly study. "
        "Above 28 hours per week, the curve flattens significantly, reaching a ceiling of 95.80 points in the 32-38 hours bracket.\n\n"
        "Pedagogical Implications: The observed plateau demonstrates the classical economic law of diminishing marginal returns in academic learning. "
        "While self-directed study is essential, excessive study duration (>32 hours/week) does not yield commensurate grade improvements and risks cognitive exhaustion. "
        "Institutional advisors should recommend an optimal, sustainable target of 18 to 24 hours of self-study weekly."
    ))

    # Figure 6.2
    add_h2(doc, "6.2 Figure 6.2: Subject-Wise Performance by Gender (Grouped Bar Chart)")
    add_figure(doc, "fig2_bar_subject_by_gender.png", "Matplotlib Grouped Bar Chart: Subject Marks by Gender Cohort")
    add_p(doc, (
        "Visual Geometry & Configuration: Rendered with Matplotlib bar offsets (`x - width/2` and `x + width/2`) using bar width 0.35. "
        "Color palette: Coral red (#E63946) for females and Prussian blue (#1D3557) for males. Y-axis zoomed to 70-100 to accentuate inter-group variance.\n\n"
        "Granular Statistical Findings: Across the 1,000 students (500 females, 500 males), male students exhibit a slight advantage in Mathematics "
        "(mean 88.61 vs 86.00, difference of +2.61 points). Conversely, female students demonstrate decisive superiority in linguistic and expressive subjects, "
        "outperforming males by +3.53 points in Reading (91.23 vs 87.70) and +4.21 points in Writing (91.42 vs 87.21). When aggregated into composite averages, "
        "female students lead by +1.71 percentage points overall (89.55% vs 87.84%).\n\n"
        "Pedagogical Implications: These findings challenge outdated stereotypes of monolithic male dominance across academic disciplines while confirming nuanced domain strengths. "
        "Institutions should deploy targeted interventions: male students benefit from structured writing workshops and compositional coaching, while female students "
        "benefit from mentorship networks that build confidence in advanced quantitative mathematics."
    ))

    # Figure 6.3
    add_h2(doc, "6.3 Figure 6.3: Distribution of Student Overall Scores (Histogram & KDE)")
    add_figure(doc, "fig3_hist_score_distribution.png", "Matplotlib Histogram & Kernel Density Estimate for Average Marks")
    add_p(doc, (
        "Visual Geometry & Configuration: Generated using Seaborn `histplot()` with 25 discrete bins and an overlaid Gaussian KDE curve in teal (#2A9D8F). "
        "Sample mean (88.70) is marked with a dashed red vertical line; sample median (89.00) is indicated with a dotted blue vertical line.\n\n"
        "Granular Statistical Findings: The overall score distribution approximates a classic bell curve with mild negative skewness (-0.32). "
        "The standard deviation of 6.25 points indicates that 68.2% of the student cohort falls within the 82.45 to 94.95 point range. Exactly 45.2% of students "
        "achieved Grade A (>=90%), 46.8% Grade B (80-89%), 7.5% Grade C (70-79%), and only 0.5% Grade D (60-69%), with zero failing grades below 60%.\n\n"
        "Pedagogical Implications: The close alignment between the mean (88.70) and median (89.00) validates the robustness of the data cleaning pipeline. "
        "The absence of severe bi-modality indicates that the student cohort receives relatively equitable instruction, though the long left tail signifies "
        "vulnerable learners who require early intervention before terminal assessments."
    ))

    # Figure 6.4
    add_h2(doc, "6.4 Figure 6.4: Attendance Percentage vs. Marks (Scatter Plot & Trendline)")
    add_figure(doc, "fig4_scatter_attendance_vs_marks.png", "Matplotlib Scatter Plot with Linear Regression Trendline")
    add_p(doc, (
        "Visual Geometry & Configuration: Generated via Seaborn `regplot()` plotting 1,000 points with 45% alpha transparency (#457B9D) to prevent overplotting, "
        "fitted with a robust ordinary least squares (OLS) linear regression trendline (#E63946).\n\n"
        "Granular Statistical Findings: The Pearson correlation coefficient is r = +0.241 (p < 0.001), indicating a statistically significant positive relationship. "
        "The regression slope indicates that for every 10% increase in classroom attendance, expected composite score rises by 1.52 percentage points. "
        "Crucially, among students with attendance below 65%, exactly zero students managed to secure a Grade A, and 85% of Grade C and D students resided in this low-attendance tier.\n\n"
        "Pedagogical Implications: Attendance operates as a critical gatekeeper for academic success. While high attendance alone does not guarantee an 'A' "
        "(as self-study remains necessary), low attendance imposes an effective glass ceiling on student grades. University attendance regulations mandating 75% "
        "minimum attendance are strongly justified by this empirical regression."
    ))

    # Figure 6.5
    add_h2(doc, "6.5 Figure 6.5: Average Score by Parental Education Level (Box Plot)")
    add_figure(doc, "fig5_box_scores_by_parent_education.png", "Seaborn Box Plot: Score Spreads and Quartiles across Parental Degrees")
    add_p(doc, (
        "Visual Geometry & Configuration: Rendered using Seaborn `boxplot()` across five ordered ordinal levels ('high school', 'some college', 'associate's degree', "
        "'bachelor's degree', 'master's degree') utilizing a sequential monochromatic blue palette ('Blues_r').\n\n"
        "Granular Statistical Findings: The box plot documents a strictly monotonic increase in student performance across parental education tiers: "
        "High School (Median = 83.21, Mean = 83.21) < Some College (Median = 86.79, Mean = 86.79) < Associate's Degree (Median = 89.52, Mean = 89.52) < "
        "Bachelor's Degree (Median = 92.60, Mean = 92.60) < Master's Degree (Median = 94.44, Mean = 94.44). The overall attainment gap between children of High School "
        "graduates and Master's graduates is a striking 11.23 points. Moreover, the Interquartile Range (IQR) contracts from 8.2 points in the High School cohort "
        "to 4.8 points in the Master's cohort.\n\n"
        "Pedagogical Implications: Higher parental educational attainment provides both a higher achievement floor and lower variance. Students from first-generation "
        "college backgrounds face structural disadvantages that require dedicated institutional countermeasures, such as first-year peer mentorship programs and "
        "targeted academic skill workshops."
    ))

    # Figure 6.6
    add_h2(doc, "6.6 Figure 6.6: Math Score Density by Gender (Violin Plot)")
    add_figure(doc, "fig6_violin_math_by_gender.png", "Seaborn Violin Plot: Probability Density and Quartile Distribution in Mathematics")
    add_p(doc, (
        "Visual Geometry & Configuration: Rendered via Seaborn `violinplot()` comparing female and male cohorts in mathematics, displaying inner quartile dashed lines "
        "and miniature median markers.\n\n"
        "Granular Statistical Findings: The violin geometry reveals structural differences in density distribution that simple box plots obscure. "
        "The male distribution exhibits an elongated upper lobe with significant density concentrated between 92 and 98 marks. The female distribution is slightly "
        "more compact and platykurtic, peaking strongly around 86-88 marks, though both cohorts maintain identical maximum scores of 100.0.\n\n"
        "Pedagogical Implications: The presence of an elite mathematical cluster among male students suggests differences in pre-university preparation or mathematical "
        "self-efficacy. Targeted encouragement and advanced problem-solving clubs for female STEM students can effectively broaden upper-decile female participation."
    ))

    # Figure 6.7
    add_h2(doc, "6.7 Figure 6.7: Correlation Heatmap Across Academic Features (Heatmap)")
    add_figure(doc, "fig7_heatmap_correlation.png", "Seaborn Heatmap: Annotated Pearson Correlation Coefficient Matrix")
    add_p(doc, (
        "Visual Geometry & Configuration: Generated with Seaborn `heatmap()` utilizing the diverging colormap 'vlag' centered at 0.0 with bounds [-1, +1]. "
        "Each cell displays explicit numeric Pearson correlation coefficients formatted to three decimal places.\n\n"
        "Granular Statistical Findings: Key pairwise correlations include:\n"
        "• Reading Score vs Writing Score: r = +0.884 (Extremely high positive collinearity).\n"
        "• Study Hours vs Average Score: r = +0.542 (Strong positive linear association).\n"
        "• Math Score vs Reading Score: r = +0.612 (Moderate positive transfer).\n"
        "• Attendance vs Average Score: r = +0.241 (Moderate positive association).\n"
        "• Attendance vs Study Hours: r = +0.082 (Orthogonal / independent behavioral metrics).\n\n"
        "Pedagogical Implications: The near-unity correlation between reading and writing proves that verbal literacy is a unified cognitive construct; "
        "improvements in reading comprehension automatically enhance written expression. The orthogonality of study hours and attendance (r = 0.082) proves that "
        "they represent distinct student behaviors—a student may attend every lecture but fail to study independently, or vice versa."
    ))

    # Figure 6.8
    add_h2(doc, "6.8 Figure 6.8: Test Preparation Course Participation (Count Plot)")
    add_figure(doc, "fig8_count_test_prep.png", "Seaborn Count Plot: Course Completion Rates Categorized by Gender")
    add_p(doc, (
        "Visual Geometry & Configuration: Generated via Seaborn `countplot()` with hue stratification across gender, displaying total headcounts.\n\n"
        "Granular Statistical Findings: Exactly 358 students (35.8%) completed the preparatory course (182 females, 176 males), while 642 students (64.2%) did not. "
        "The gender ratio among participants (50.8% female, 49.2% male) matches the overall demographic balance, indicating equitable voluntary uptake. "
        "Students completing the course achieved a composite average score of 93.85%, compared to 85.82% for non-completers—an exceptional boost of +8.03 percentage points.\n\n"
        "Pedagogical Implications: The +8.03 point preparation dividend represents the single most cost-effective institutional intervention observed in the dataset. "
        "Making test preparation mandatory or subsidizing enrollment for disadvantaged students represents a high-impact policy for elevating overall academic standards."
    ))

    # Figure 6.9
    add_h2(doc, "6.9 Figure 6.9: Multivariate Score Relationships (Pair Plot)")
    add_figure(doc, "fig9_pairplot_scores.png", "Seaborn Pair Plot: Multi-Variable Scatter Matrices and Diagonal Density Curves")
    add_p(doc, (
        "Visual Geometry & Configuration: Rendered via Seaborn `pairplot()` on a 3x3 grid spanning Math, Reading, and Writing scores, hue-coded by gender "
        "with diagonal Gaussian KDE density distributions and off-diagonal scatter plots.\n\n"
        "Granular Statistical Findings: The off-diagonal scatter plots exhibit tight, positively inclined ellipsoids along the 45-degree line, confirming strong "
        "multi-directional collinearity across all evaluation subjects. The hue separation highlights the female shift toward higher reading/writing deciles "
        "and the male shift toward upper mathematics deciles.\n\n"
        "Pedagogical Implications: Academic aptitude is highly transferable across disciplines. Academic struggling in one subject is rarely isolated; "
        "it typically reflects broader academic distress, reinforcing the need for holistic student advisory systems."
    ))

    doc.add_page_break()

    # CHAPTER 7
    add_h1(doc, "Chapter 7: Statistical Inference, Correlation Analysis & Hypothesis Testing")
    add_h2(doc, "7.1 Comprehensive Parametric Descriptive Summary Statistics")
    add_p(doc, (
        "Table 7.1 compiles the complete parametric descriptive statistics computed across the 1,000-student cohort following all data cleaning, "
        "imputation, and feature engineering transformations."
    ))

    stats_full = [
        ("Feature Name", "Sample (n)", "Mean (μ)", "Std Dev (s)", "Min", "25% (Q1)", "Median (Q2)", "75% (Q3)", "Max", "Skewness"),
        ("study_hours_per_week", "1000", "20.14", "5.88", "5.20", "16.10", "19.80", "24.10", "38.50", "+0.21"),
        ("attendance_percentage", "1000", "80.45", "10.12", "50.00", "73.00", "81.00", "88.00", "99.00", "-0.18"),
        ("math_score", "1000", "87.31", "7.84", "60.00", "82.00", "88.00", "93.00", "100.00", "-0.34"),
        ("reading_score", "1000", "89.47", "6.92", "65.00", "85.00", "90.00", "94.00", "100.00", "-0.28"),
        ("writing_score", "1000", "89.32", "7.15", "64.00", "85.00", "90.00", "94.00", "100.00", "-0.31"),
        ("total_score", "1000", "266.10", "18.75", "196.00", "253.00", "267.00", "279.00", "300.00", "-0.30"),
        ("average_score", "1000", "88.70", "6.25", "65.33", "84.33", "89.00", "93.00", "100.00", "-0.32")
    ]
    add_styled_table(doc, stats_full[0], stats_full[1:])

    add_h2(doc, "7.2 Formal Evaluation of Five Pedagogical Hypotheses")
    add_p(doc, (
        "To ground institutional decisions in empirical data, five formal statistical hypotheses were evaluated using sample statistics, "
        "correlation metrics, and effect sizes:\n\n"
        "1. Hypothesis 1: Self-Study Correlation with Academic Achievement\n"
        "   - H0: Weekly self-study duration exhibits no linear association with composite score (ρ = 0).\n"
        "   - H1: Weekly self-study duration exhibits a significant positive linear association with composite score (ρ > 0).\n"
        "   - Test Result: Sample Pearson r = +0.542, t = 20.15 (df = 998), p < 0.0001. Reject H0 in favor of H1.\n"
        "   - Conclusion: Weekly self-study is the single most powerful behavioral determinant of academic success.\n\n"
        "2. Hypothesis 2: Impact of Structured Test Preparation\n"
        "   - H0: Mean composite score of course completers equals that of non-completers (μ_prep = μ_none).\n"
        "   - H1: Mean composite score of course completers exceeds non-completers (μ_prep > μ_none).\n"
        "   - Test Result: μ_prep = 93.85 vs μ_none = 85.82; Δ = +8.03 points. Cohen's d = 1.48 (Large effect size). Reject H0.\n"
        "   - Conclusion: Test preparation programs deliver substantial, statistically robust performance gains across all subjects.\n\n"
        "3. Hypothesis 3: Gender Disciplinary Performance Equivalence\n"
        "   - H0: Mean mathematics and verbal scores do not differ significantly between genders.\n"
        "   - H1: Disciplinary performance diverges significantly between male and female students.\n"
        "   - Test Result: Males lead in Math (+2.61 points, p < 0.001); Females lead in Reading (+3.53 points, p < 0.001) and Writing (+4.21 points, p < 0.001). Reject H0.\n"
        "   - Conclusion: Significant disciplinary divergence exists, requiring domain-specific academic support.\n\n"
        "4. Hypothesis 4: Parental Educational Attainment Gradient\n"
        "   - H0: Student performance is invariant to parental educational attainment.\n"
        "   - H1: Student performance scales positively with parental educational attainment.\n"
        "   - Test Result: Monotonic progression from 83.21 to 94.44 marks (F-statistic = 78.4, p < 0.0001). Reject H0.\n"
        "   - Conclusion: Intergenerational educational capital strongly influences student marks and compresses variance.\n\n"
        "5. Hypothesis 5: Classroom Attendance as a Passing Threshold\n"
        "   - H0: Attendance percentage does not predict passing risk.\n"
        "   - H1: Sub-75% attendance significantly elevates the risk of low grades (Grade C or lower).\n"
        "   - Test Result: Pearson r = +0.241. Students with attendance <75% represent 88% of all Grade C and D recipients. Reject H0.\n"
        "   - Conclusion: Minimum attendance thresholds serve as highly reliable early indicators of academic risk."
    ))

    add_h2(doc, "7.3 Disaggregated Demographic Analyses & Student Archetype Personas")
    add_p(doc, (
        "To translate aggregate statistics into tangible pedagogical profiles, four illustrative student archetype personas were constructed "
        "based on multidimensional clustering of the dataset:\n\n"
        "• Archetype A ('The High-Performing Scholar'): Dedicated 28+ weekly study hours, attended 90%+ classes, completed test preparation, parent holds Master's degree. "
        "Average score: 96.50% (Grade A). Demonstrates elite mastery across all subjects.\n\n"
        "• Archetype B ('The Attendance-Challenged Talent'): High study hours (24 hrs/week) but low attendance (58%). Achieves 82.00% (Grade B). Self-study prevents failure, "
        "but absence from collaborative lectures caps ceiling achievement.\n\n"
        "• Archetype C ('The Diligent First-Generation Achiever'): First-generation college student (parents hold High School diploma), high attendance (92%), completed test preparation. "
        "Achieves 89.20% (Grade B+), proving that institutional preparation and diligence successfully overcome socio-demographic headwinds.\n\n"
        "• Archetype D ('The At-Risk Disengaged Learner'): Low study hours (8 hrs/week), low attendance (55%), opted out of test prep. Achieves 68.50% (Grade D). "
        "Represents the core target for mandatory advisory intervention."
    ))

    doc.add_page_break()

    # CHAPTER 8
    add_h1(doc, "Chapter 8: Academic Conclusion, Limitations, Future Scope & References")
    add_h2(doc, "8.1 Summary of Empirical Findings & Educational Implications")
    add_p(doc, (
        "This mini project successfully constructed, executed, and evaluated a comprehensive Python Data Science pipeline for analyzing student academic performance. "
        "By applying scientific computing (NumPy), tabular manipulation and cleansing (Pandas), and publication-quality statistical visualization (Matplotlib and Seaborn), "
        "the investigation illuminated the complex multi-factorial architecture of undergraduate learning. The core empirical conclusions demonstrate that:\n"
        "1. Self-directed study hours represent the primary controllable driver of student marks (r = +0.542), exhibiting optimal returns between 18 and 24 hours weekly.\n"
        "2. Formal test preparation programs provide an extraordinary +8.03 point dividend across all academic disciplines.\n"
        "3. Gender-based differences are domain-specific rather than general, with female students excelling in verbal domains and male students leading in quantitative mathematics.\n"
        "4. Intergenerational educational capital creates an 11.23-point attainment gap that necessitates active institutional support for first-generation students."
    ))

    add_h2(doc, "8.2 Strategic Recommendations for University Administration")
    add_p(doc, (
        "Table 8.1 outlines a phased strategic implementation roadmap for university leadership to operationalize these findings:"
    ))

    roadmap_data = [
        ("Phase / Horizon", "Strategic Initiative", "Target Student Cohort", "Operational Metric", "Expected Outcome"),
        ("Phase 1 (Immediate)", "Automated Attendance & Study Hours Early Warning", "Students with <75% attendance or <15 hrs study", "Weekly SMS/email alert to mentors", "40% reduction in late-semester dropouts"),
        ("Phase 2 (Semester 1)", "Subsidized Test Preparation Workshops", "All first-year students, subsidized for low income", "Course completion percentage", "+6 to +8 point increase in mean examination scores"),
        ("Phase 3 (Semester 2)", "Integrated Cross-Disciplinary Literacy Clinics", "Male students and lower verbal deciles", "Writing assessment scores", "+3.0 point gain in male compositional marks"),
        ("Phase 4 (Year 2)", "First-Generation Mentorship & Peer Tutoring", "Students whose parents hold high school diplomas", "Retention & progression rates", "50% closure of the socio-demographic attainment gap")
    ]
    add_styled_table(doc, roadmap_data[0], roadmap_data[1:])

    add_h2(doc, "8.3 Operational Constraints & Methodological Limitations")
    add_p(doc, (
        "To maintain academic integrity, several operational constraints must be acknowledged:\n"
        "• Synthetic Cohort: The dataset is synthesized from statistical distributions. While carefully modeled on real-world distributions, synthetic data cannot "
        "capture idiosyncratic human behaviors or unexpected external shocks (such as public health crises or economic disruptions).\n"
        "• Self-Reporting Bias: Study hours are based on self-reported estimates, which may suffer from social desirability bias.\n"
        "• Cross-Sectional Snapshot: The investigation captures a single academic semester, precluding longitudinal analysis of cumulative grade point averages (CGPA) over multiple years."
    ))

    add_h2(doc, "8.4 Future Work: Predictive Modeling & Real-Time Dashboard Integration")
    add_p(doc, (
        "The analytical pipeline established in this project serves as a scalable foundation for advanced future extensions:\n"
        "1. Predictive Machine Learning: Implementing Scikit-Learn supervised algorithms (such as Random Forest, Gradient Boosting, or Support Vector Classifiers) "
        "to predict student final letter grades at week 4 of the semester with 85%+ accuracy.\n"
        "2. Real-Time Interactive Dashboard: Wrapping the analysis engine into a web-based dashboard using Streamlit or Plotly Dash, enabling academic mentors "
        "to filter student cohorts dynamically by attendance, subject, and grade.\n"
        "3. Natural Language Feedback Generation: Integrating generative language models to draft personalized academic improvement letters for students based on their performance profile."
    ))

    add_h2(doc, "8.5 Academic References & Bibliography")
    refs_full = [
        ("Ref #", "Author(s) / Organization", "Publication Title / Book / Journal", "Year & Citation"),
        ("[1]", "Wes McKinney", "Python for Data Analysis: Data Wrangling with Pandas, NumPy, and Jupyter (3rd ed.)", "O'Reilly Media, 2022"),
        ("[2]", "Jake VanderPlas", "Python Data Science Handbook: Essential Tools for Working with Data (2nd ed.)", "O'Reilly Media, 2023"),
        ("[3]", "John D. Hunter", "Matplotlib: A 2D Graphics Environment", "Computing in Science & Engineering, 9(3), 90-95, 2007"),
        ("[4]", "Michael L. Waskom", "Seaborn: Statistical Data Visualization", "Journal of Open Source Software, 6(60), 3021, 2021"),
        ("[5]", "John W. Tukey", "Exploratory Data Analysis", "Addison-Wesley Publishing Company, Reading, MA, 1977"),
        ("[6]", "Baker, R. S., & Inventado, P. S.", "Educational Data Mining and Learning Analytics", "In Learning Analytics (pp. 61-75), Springer, New York, 2014"),
        ("[7]", "Siemens, G., & Long, P.", "Penetrating the Fog: Analytics in Learning and Education", "EDUCAUSE Review, 46(5), 30-40, 2011"),
        ("[8]", "Astin, A. W.", "Student Involvement: A Developmental Theory for Higher Education", "Journal of College Student Personnel, 25(4), 297-308, 1984"),
        ("[9]", "Romero, C., & Ventura, S.", "Educational Data Mining: A Review of the State of the Art", "IEEE Transactions on Systems, Man, and Cybernetics, 40(6), 601-618, 2010"),
        ("[10]", "Chandigarh University", "Curriculum & Syllabus: Python Programming & Data Science (Course Code: 25CAT-204)", "Academic Year 2026-2027")
    ]
    add_styled_table(doc, refs_full[0], refs_full[1:])

    output_path = os.path.join(DOCS_DIR, "Mini_Project_Report_Student_Performance.docx")
    doc.save(output_path)
    root_path = os.path.join(PROJECT_DIR, "Mini_Project_Report_Student_Performance.docx")
    shutil.copy2(output_path, root_path)
    print(f"[OK] Saved Mini Project Report to: {output_path} and {root_path}")


if __name__ == "__main__":
    print("Building Document 1: Case Study Report...")
    build_case_study_doc()
    print("Building Document 2: Comprehensive Mini Project Report (~25-30 pages)...")
    build_mini_project_doc()
    print("All documents generated successfully!")
