import pandas as pd

def calculate_level(total):
    if total >= 85:
        return '5**'
    elif total >= 77:
        return '5*'
    elif total >= 70:
        return '5'
    elif total >= 55:
        return '4'
    elif total >= 45:
        return '3'
    elif total >= 30:
        return '2'
    elif total >= 20:
        return '1'
    elif total >= 5:
        return 'U'
    else:
        return 'Error'  # Error catch

def process_student_scores(df):
    weight_paper_i = 0.6875
    weight_paper_ii = 0.3125

    df['total_score'] = 0
    df['level'] = ''
    
    for index, row in df.rows():
        # Summing scores for Paper I and Paper II according to the given formula
        paper_i = sum([row[f'MC'], row['Bq1'], row['Bq2'], row['Bq3'], row['Bq4'], row['Bq5']])
        paper_ii = sum([row['Eq1'], row['Eq2'], row['Eq3'], row['Eq4']])
        
        total_weighted_score = (paper_i * weight_paper_i) + (paper_ii * weight_paper_ii)
        level = calculate_level(total_weighted_score)
        
        df.at[index, 'total_score'] = total_weighted_score
        df.at[index, 'level'] = level

    return df