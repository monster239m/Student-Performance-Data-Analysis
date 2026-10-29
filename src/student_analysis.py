import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def run_analysis():
    # Set plot aesthetic
    sns.set_theme(style="whitegrid", palette="muted")
    plt.rcParams.update({'font.sans-serif': 'Arial', 'font.size': 10})

    csv_path = r"C:\Users\Rahul\.gemini\antigravity\scratch\sem-project\python-student-performance\data\student_performance.csv"
    fig_dir = r"C:\Users\Rahul\.gemini\antigravity\scratch\sem-project\python-student-performance\figures"
    os.makedirs(fig_dir, exist_ok=True)

    print("=== STEP 1: LOADING DATASET ===")
    df = pd.read_csv(csv_path)
    print("Initial Shape:", df.shape)

    print("\n=== STEP 2: DATA CLEANING USING PANDAS ===")
    # 1. Duplicates
    dup_count = df.duplicated().sum()
    print(f"Duplicate rows detected: {dup_count}")
    df = df.drop_duplicates().reset_index(drop=True)

    # 2. String Standardization
    df['gender'] = df['gender'].str.capitalize()

    # 3. Missing Values
    null_counts = df.isnull().sum()
    print("Missing values per column:\n", null_counts[null_counts > 0])
    
    # Impute numerical missing values with median
    if df['attendance_percent'].isnull().sum() > 0:
        df['attendance_percent'] = df['attendance_percent'].fillna(df['attendance_percent'].median())
    if df['study_hours_weekly'].isnull().sum() > 0:
        df['study_hours_weekly'] = df['study_hours_weekly'].fillna(df['study_hours_weekly'].median())
    if df['reading_score'].isnull().sum() > 0:
        df['reading_score'] = df['reading_score'].fillna(df['reading_score'].median()).astype(int)

    print("Cleaned Shape:", df.shape)
    print("Post-cleaning Nulls remaining:", df.isnull().sum().sum())

    # 4. Feature Engineering
    df['total_score'] = df['math_score'] + df['reading_score'] + df['writing_score']
    df['average_score'] = (df['total_score'] / 3.0).round(2)

    def assign_grade(avg):
        if avg >= 90: return 'A+'
        elif avg >= 80: return 'A'
        elif avg >= 70: return 'B'
        elif avg >= 60: return 'C'
        elif avg >= 50: return 'D'
        else: return 'F'

    df['grade'] = df['average_score'].apply(assign_grade)

    print("\n=== STEP 3: STATISTICAL ANALYSIS (NUMPY & PANDAS) ===")
    stats = df[['math_score', 'reading_score', 'writing_score', 'average_score', 'attendance_percent', 'study_hours_weekly']].describe().round(2)
    print(stats)

    # Groupby Gender
    gender_grp = df.groupby('gender')[['math_score', 'reading_score', 'writing_score', 'average_score']].mean().round(2)
    print("\nGender Averages:\n", gender_grp)

    # Groupby Parental Education
    edu_grp = df.groupby('parental_education')[['average_score', 'math_score', 'reading_score']].mean().round(2)
    print("\nParental Education Averages:\n", edu_grp)

    # Correlation Matrix
    corr = df[['study_hours_weekly', 'attendance_percent', 'math_score', 'reading_score', 'writing_score', 'average_score']].corr().round(3)
    print("\nCorrelation Matrix:\n", corr)

    print("\n=== STEP 4: GENERATING 9 REQUIRED VISUALIZATIONS ===")

    # 1. Matplotlib Line Chart: Study Hours vs Average Score
    plt.figure(figsize=(8, 4.5))
    binned_study = df.groupby(pd.cut(df['study_hours_weekly'], bins=8), observed=False)['average_score'].mean()
    x_labels = [f"{int(b.left)}-{int(b.right)} hrs" for b in binned_study.index]
    plt.plot(x_labels, binned_study.values, marker='o', color='#1B365D', linewidth=2.5, markersize=7)
    plt.title('Mean Student Score by Weekly Study Hours (Matplotlib Line Chart)', fontsize=12, fontweight='bold', pad=12)
    plt.xlabel('Weekly Study Hours Range', fontweight='bold')
    plt.ylabel('Mean Average Score (out of 100)', fontweight='bold')
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "fig1_line_study_vs_marks.png"), dpi=200)
    plt.close()

    # 2. Matplotlib Bar Chart: Subject Scores by Gender
    plt.figure(figsize=(8, 4.5))
    x = np.arange(3)
    width = 0.35
    female_means = [df[df['gender']=='Female']['math_score'].mean(),
                    df[df['gender']=='Female']['reading_score'].mean(),
                    df[df['gender']=='Female']['writing_score'].mean()]
    male_means = [df[df['gender']=='Male']['math_score'].mean(),
                  df[df['gender']=='Male']['reading_score'].mean(),
                  df[df['gender']=='Male']['writing_score'].mean()]
    plt.bar(x - width/2, female_means, width, label='Female', color='#2B4C7E', edgecolor='black')
    plt.bar(x + width/2, male_means, width, label='Male', color='#E67E22', edgecolor='black')
    plt.ylabel('Average Score', fontweight='bold')
    plt.title('Subject-wise Performance Comparison by Gender (Matplotlib Bar Chart)', fontsize=12, fontweight='bold', pad=12)
    plt.xticks(x, ['Math Score', 'Reading Score', 'Writing Score'], fontweight='bold')
    plt.ylim(0, 100)
    plt.legend(frameon=True)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "fig2_bar_subject_by_gender.png"), dpi=200)
    plt.close()

    # 3. Matplotlib Histogram: Overall Score Distribution
    plt.figure(figsize=(8, 4.5))
    n_counts, bins, patches = plt.hist(df['average_score'], bins=20, color='#3498DB', edgecolor='black', alpha=0.85)
    plt.title('Distribution of Student Average Scores (Matplotlib Histogram)', fontsize=12, fontweight='bold', pad=12)
    plt.xlabel('Average Score', fontweight='bold')
    plt.ylabel('Frequency (Number of Students)', fontweight='bold')
    plt.axvline(df['average_score'].mean(), color='red', linestyle='dashed', linewidth=2, label=f'Mean: {df["average_score"].mean():.1f}')
    plt.axvline(df['average_score'].median(), color='green', linestyle='dashed', linewidth=2, label=f'Median: {df["average_score"].median():.1f}')
    plt.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "fig3_hist_score_distribution.png"), dpi=200)
    plt.close()

    # 4. Matplotlib Scatter Plot: Attendance vs Average Score
    plt.figure(figsize=(8, 4.5))
    plt.scatter(df['attendance_percent'], df['average_score'], alpha=0.6, c=df['average_score'], cmap='viridis', edgecolors='w', s=45)
    m, b = np.polyfit(df['attendance_percent'], df['average_score'], 1)
    plt.plot(df['attendance_percent'], m*df['attendance_percent'] + b, color='red', linewidth=2, label=f'Trendline (Slope = +{m:.2f})')
    plt.title('Attendance Percentage vs Average Academic Score (Matplotlib Scatter Plot)', fontsize=12, fontweight='bold', pad=12)
    plt.xlabel('Attendance Percentage (%)', fontweight='bold')
    plt.ylabel('Average Academic Score', fontweight='bold')
    plt.colorbar(label='Score Level')
    plt.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "fig4_scatter_attendance_vs_marks.png"), dpi=200)
    plt.close()

    # 5. Seaborn Box Plot: Scores across Parental Education Levels
    plt.figure(figsize=(9, 4.5))
    order = ['High School', 'Some College', "Associate's Degree", "Bachelor's Degree", "Master's Degree"]
    sns.boxplot(x='parental_education', y='average_score', data=df, order=order, palette='Set2')
    plt.title('Academic Score Spread Across Parental Education Levels (Seaborn Box Plot)', fontsize=12, fontweight='bold', pad=12)
    plt.xlabel('Parental Education Level', fontweight='bold')
    plt.ylabel('Average Score', fontweight='bold')
    plt.xticks(rotation=15)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "fig5_box_scores_by_parent_education.png"), dpi=200)
    plt.close()

    # 6. Seaborn Violin Plot: Math Score Density by Gender
    plt.figure(figsize=(8, 4.5))
    sns.violinplot(x='gender', y='math_score', data=df, palette='Blues', inner='quartile')
    plt.title('Math Score Probability Density by Gender (Seaborn Violin Plot)', fontsize=12, fontweight='bold', pad=12)
    plt.xlabel('Gender', fontweight='bold')
    plt.ylabel('Math Score', fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "fig6_violin_math_by_gender.png"), dpi=200)
    plt.close()

    # 7. Seaborn Heatmap: Correlation Matrix
    plt.figure(figsize=(7, 5))
    sns.heatmap(corr, annot=True, cmap='Blues', fmt='.2f', linewidths=0.5, cbar_kws={'label': 'Pearson Correlation (r)'})
    plt.title('Feature Correlation Heatmap (Seaborn Heatmap)', fontsize=12, fontweight='bold', pad=12)
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "fig7_heatmap_correlation.png"), dpi=200)
    plt.close()

    # 8. Seaborn Count Plot: Test Prep Course Participation by Gender
    plt.figure(figsize=(8, 4.5))
    sns.countplot(x='test_prep_course', hue='gender', data=df, palette=['#3498DB', '#E67E22'])
    plt.title('Test Preparation Course Enrollment by Gender (Seaborn Count Plot)', fontsize=12, fontweight='bold', pad=12)
    plt.xlabel('Test Preparation Status', fontweight='bold')
    plt.ylabel('Student Count', fontweight='bold')
    plt.legend(title='Gender')
    plt.tight_layout()
    plt.savefig(os.path.join(fig_dir, "fig8_count_test_prep.png"), dpi=200)
    plt.close()

    # 9. Seaborn Pair Plot: Multivariate Score & Attendance Relationships
    pair_cols = ['math_score', 'reading_score', 'writing_score', 'attendance_percent']
    g = sns.pairplot(df[pair_cols], diag_kind='kde', plot_kws={'alpha': 0.5, 's': 20, 'color': '#1B365D'}, diag_kws={'color': '#E67E22'})
    g.fig.subplots_adjust(top=0.94)
    g.fig.suptitle('Multivariate Pairwise Score Distributions (Seaborn Pair Plot)', fontsize=13, fontweight='bold')
    g.savefig(os.path.join(fig_dir, "fig9_pairplot_scores.png"), dpi=180)
    plt.close()

    print("All 9 figures generated successfully in:", fig_dir)

if __name__ == '__main__':
    run_analysis()
