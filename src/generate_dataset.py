import numpy as np
import pandas as pd

def generate_student_data():
    np.random.seed(42)
    n = 1000

    # Categorical distributions
    genders = np.random.choice(['Female', 'Male'], size=n, p=[0.52, 0.48])
    parental_edu = np.random.choice(
        ['High School', 'Some College', "Associate's Degree", "Bachelor's Degree", "Master's Degree"],
        size=n,
        p=[0.20, 0.25, 0.25, 0.20, 0.10]
    )
    lunch = np.random.choice(['Standard', 'Free/Reduced'], size=n, p=[0.65, 0.35])
    test_prep = np.random.choice(['Completed', 'None'], size=n, p=[0.38, 0.62])

    # Numerical features
    study_hours = np.clip(np.random.normal(15, 6, size=n), 2.0, 38.0).round(1)
    attendance = np.clip(np.random.normal(82, 10, size=n), 45.0, 100.0).round(1)

    # Base academic capability
    edu_boost = {
        'High School': 0,
        'Some College': 4,
        "Associate's Degree": 7,
        "Bachelor's Degree": 11,
        "Master's Degree": 15
    }
    boost_val = np.array([edu_boost[e] for e in parental_edu])
    prep_val = np.array([8 if t == 'Completed' else 0 for t in test_prep])
    lunch_val = np.array([5 if l == 'Standard' else -3 for l in lunch])

    # Core scores with realistic correlations
    base = 40 + (study_hours * 0.9) + (attendance * 0.25) + boost_val + prep_val + lunch_val
    math_scores = np.clip(base + np.random.normal(2, 7, size=n) + (genders == 'Male') * 3, 20, 100).astype(int)
    reading_scores = np.clip(base + np.random.normal(4, 6, size=n) + (genders == 'Female') * 4, 25, 100).astype(int)
    writing_scores = np.clip(base + np.random.normal(3, 6, size=n) + (genders == 'Female') * 5, 25, 100).astype(int)

    df = pd.DataFrame({
        'student_id': [f'STU_{1001 + i}' for i in range(n)],
        'gender': genders,
        'parental_education': parental_edu,
        'lunch_type': lunch,
        'test_prep_course': test_prep,
        'study_hours_weekly': study_hours,
        'attendance_percent': attendance,
        'math_score': math_scores,
        'reading_score': reading_scores,
        'writing_score': writing_scores
    })

    # Intentionally inject 5 null values, 3 duplicates, and 2 inconsistent values for data cleaning demonstration
    df.loc[12, 'attendance_percent'] = np.nan
    df.loc[45, 'attendance_percent'] = np.nan
    df.loc[88, 'study_hours_weekly'] = np.nan
    df.loc[120, 'reading_score'] = np.nan
    df.loc[150, 'gender'] = 'female' # lowercase inconsistency
    df.loc[210, 'gender'] = 'MALE'   # uppercase inconsistency

    # Append 3 duplicate rows
    dups = df.iloc[[10, 25, 50]].copy()
    df = pd.concat([df, dups], ignore_index=True)

    out_csv = r"C:\Users\Rahul\.gemini\antigravity\scratch\sem-project\python-student-performance\data\student_performance.csv"
    df.to_csv(out_csv, index=False)
    print("Dataset generated successfully:", out_csv, "with shape:", df.shape)

if __name__ == '__main__':
    generate_student_data()
