import pandas as pd
import os
# Assuming calculate_scores and calculate_level functions are correctly imported
from data_processing.calculate_scores import calculate_level
from database.connect import insert_score
from data_processing.calculate_scores import process_student_scores # Ensure you have this function.
from data_processing.validation import validate_df
def import_and_process_file(file_path):
    school_id = os.path.splitext(os.path.basename(file_path))[0]

    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        raise ValueError("Only .xlsx and .csv are supported.")

    validate_df(df)  # Add this line to check if the DataFrame structure is correct.

    for index, row in df.iterrows():
        scores_dict = {
            'MC': row['Mc'],
            'Bq': [row[f'Bq{i}'] for i in range(1, 6)],
            'Eq': [row[f'Eq{i}'] for i in range(1, 5)]
        }
        
        total_score_paper_i, total_score_paper_ii = process_student_scores(scores_dict)
        
        total_weighted_score = total_score_paper_i * 0.6875 + total_score_paper_ii * 0.3125
        level = calculate_level(total_weighted_score)
        
        insert_score(
            school_id, row['Classnum'], row['Name'], row['Gender'],
            row['Elective'], row['Lang'], row['Mc'], row['Bq1'], row['Bq2'], 
            row['Bq3'], row['Bq4'], row['Bq5'], row['Eq1'], row['Eq2'], 
            row['Eq3'], row['Eq4'], total_weighted_score, level
        )
    
    print(f"Data from {file_path} processed and inserted.")