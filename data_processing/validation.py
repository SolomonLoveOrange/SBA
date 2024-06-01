def validate_df(df):
    expected_columns = [
        'Classnum', 'Name', 'Gender', 'Elective', 'Lang', 'Mc', 'Bq1', 
        'Bq2', 'Bq3', 'Bq4', 'Bq5', 'Eq1', 'Eq2', 'Eq3', 'Eq4'
    ]
    missing_columns = set(expected_columns) - set(df.columns)
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")