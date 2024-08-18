import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import sqlite3
import openpyxl
import csv
import pandas as pd
import os
import matplotlib
matplotlib.use('Agg')  # Set the backend to Agg
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def drop_and_create_database(database_name):
    try:
        conn = sqlite3.connect(database_name)
        print(f"Database '{database_name}' opened/created.")
    except sqlite3.Error as err:
        print(f"Error: {err}")
    finally:
        if conn:
            conn.close()
            print(f"Database '{database_name}' closed.")

def clear_database_table(database_name, table_name):
    try:
        conn = sqlite3.connect(database_name)
        cursor = conn.cursor()
        cursor.execute(f"DELETE FROM {table_name}")
        conn.commit()
        print(f"Table '{table_name}' cleared successfully.")
    except sqlite3.Error as err:
        print(f"Error clearing table: {err}")
    finally:
        if conn:
            conn.close()

def create_fresh_db_tables(database_name):
    try:
        conn = sqlite3.connect(database_name)
        cursor = conn.cursor()
        query = """
        CREATE TABLE IF NOT EXISTS student_computed_grade (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            School_name TEXT NOT NULL,
            Classnum VARCHAR(5) NOT NULL,
            Name VARCHAR(50) NOT NULL,
            Gender CHAR(1) NOT NULL CHECK (Gender IN ('M', 'F')),
            Elective CHAR(2) NOT NULL CHECK (Elective IN ('AB', 'AC', 'BC')),
            Lang CHAR(1) NOT NULL CHECK (Lang IN ('C', 'E')),
            Mc INT NOT NULL,
            Bq1 INT NOT NULL,
            Bq2 INT NOT NULL,
            Bq3 INT NOT NULL,
            Bq4 INT NOT NULL,
            Bq5 INT NOT NULL,
            Eq1 INT NOT NULL,
            Eq2 INT NOT NULL,
            Eq3 INT NOT NULL,
            Eq4 INT NOT NULL,
            ComputedSCORE DECIMAL(10,2) DEFAULT NULL,
            Level VARCHAR(5) DEFAULT NULL
        )
        """
        cursor.execute(query)
        print("Database table 'student_computed_grade' created successfully.")
    except sqlite3.Error as err:
        print(f"Error: {err}")
    finally:
        if conn:
            conn.close()

def validate_data(row):
    if not isinstance(row['Classnum'], str):
        raise ValueError(f"Invalid Classnum for {row['Name']}: must be a string")
    if not isinstance(row['Name'], str):
        raise ValueError(f"Invalid Name: {row['Name']}")
    if row['Gender'] not in ['M', 'F']:
        raise ValueError(f"Invalid Gender for {row['Name']}: must be 'M' or 'F'")
    if row['Elective'] not in ['AB', 'AC', 'BC']:
        raise ValueError(f"Invalid Elective for {row['Name']}: must be 'AB', 'AC', or 'BC'")
    if row['Lang'] not in ['C', 'E']:
        raise ValueError(f"Invalid Lang for {row['Name']}: must be 'C' or 'E'")
    
    if not (0 <= row['Mc'] <= 40):
        raise ValueError(f"Invalid Mc score for {row['Name']}: must be between 0 and 40")
    for bq in ['Bq1', 'Bq2', 'Bq3', 'Bq4', 'Bq5']:
        if not (0 <= row[bq] <= 12):
            raise ValueError(f"Invalid {bq} score for {row['Name']}: must be between 0 and 12")
    for eq in ['Eq1', 'Eq2', 'Eq3', 'Eq4']:
        if not (0 <= row[eq] <= 15):
            raise ValueError(f"Invalid {eq} score for {row['Name']}: must be between 0 and 15")

def import_data(folder_path, database_name, table_name):
    clear_database_table(database_name, table_name)
    try:
        conn = sqlite3.connect(database_name)
        cursor = conn.cursor()
        all_files = os.listdir(folder_path)
        data_files = [f for f in all_files if f.endswith((".xlsx", ".csv"))]
        for data_file in data_files:
            file_name, file_extension = os.path.splitext(data_file)
            filepath = os.path.join(folder_path, data_file)
           
            if file_extension == '.xlsx':
                df = pd.read_excel(filepath, engine='openpyxl')
            elif file_extension == '.csv':
                df = pd.read_csv(filepath)
           
            for _, row in df.iterrows():
                validate_data(row)
                
                values = (file_name, row['Classnum'], row['Name'], row['Gender'], row['Elective'],
                          row['Lang'], row['Mc'], row['Bq1'], row['Bq2'], row['Bq3'], row['Bq4'],
                          row['Bq5'], row['Eq1'], row['Eq2'], row['Eq3'], row['Eq4'])
                insert_query = f"""
                INSERT INTO {table_name} (School_name, Classnum, Name, Gender, Elective, Lang, Mc,
                                            Bq1, Bq2, Bq3, Bq4, Bq5, Eq1, Eq2, Eq3, Eq4)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
                cursor.execute(insert_query, values)
        conn.commit()
        print("Data import from all files completed.")
    except (sqlite3.Error, pd.errors.EmptyDataError, KeyError, ValueError) as err:
        print(f"Error: {err}")
        raise
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def compute_score(mc, bq1, bq2, bq3, bq4, bq5, eq1, eq2, eq3, eq4):
    section_a = mc * (40/40)
    section_b = (bq1/12 * 5) + (bq2/12 * 5) + (bq3/12 * 5) + (bq4/12 * 5) + (bq5/12 * 5)
    paper_i = ((section_a + section_b)/65) * 0.6875
    paper_ii = (((eq1/15 * 4) + (eq2/15 * 4) + (eq3/15 * 4) + (eq4/15 * 4))/16) * 0.3125
    total_score = paper_i + paper_ii
    percentage_score = total_score * 100
    return round(percentage_score, 2)

def determine_level(score):
    cut_scores = {
        '5**': 85, '5*': 77, '5': 70, '4': 55,
        '3': 45, '2': 30, '1': 20, 'U': 5
    }
    for level, cut_score in cut_scores.items():
        if score >= cut_score:
            return level
    return 'U'

class MarkSubmissionSystem:
    def __init__(self, master):
        self.master = master
        master.title("Mark Submission System")
        master.geometry("400x300")

        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()

        tk.Label(master, text="Input Folder:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(master, textvariable=self.input_folder, width=30).grid(row=0, column=1, padx=10, pady=10)
        tk.Button(master, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(master, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10)
        tk.Entry(master, textvariable=self.output_folder, width=30).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(master, text="Browse", command=self.browse_output).grid(row=1, column=2, padx=10, pady=10)

        tk.Button(master, text="Process Data", command=self.process_data).grid(row=2, column=1, pady=20)

        self.status_label = tk.Label(master, text="")
        self.status_label.grid(row=3, column=0, columnspan=3, pady=10)

    def browse_input(self):
        folder = filedialog.askdirectory()
        self.input_folder.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory()
        self.output_folder.set(folder)

    def process_data(self):
        if not self.input_folder.get() or not self.output_folder.get():
            messagebox.showerror("Error", "Please select both input and output folders.")
            return

        self.status_label.config(text="Processing... Please wait.")
        threading.Thread(target=self.run_processing, daemon=True).start()

    def run_processing(self):
        try:
            folder_path = self.input_folder.get()
            result_folder = self.output_folder.get()
            database_folder = os.path.join(result_folder, "database")
            os.makedirs(database_folder, exist_ok=True)
            database_name = os.path.join(database_folder, "mark_submission_system.db")
            table_name = "student_computed_grade"

            for file in os.listdir(result_folder):
                file_path = os.path.join(result_folder, file)
                if os.path.isfile(file_path):
                    os.unlink(file_path)

            drop_and_create_database(database_name)
            create_fresh_db_tables(database_name)

            import_data(folder_path, database_name, table_name)

            conn = sqlite3.connect(database_name)
            cursor = conn.cursor()

            cursor.execute("SELECT * FROM student_computed_grade")
            rows = cursor.fetchall()

            total_candidates = len(rows)
            male_candidates = sum(1 for row in rows if row[4] == 'M')
            female_candidates = sum(1 for row in rows if row[4] == 'F')
            english_candidates = sum(1 for row in rows if row[6] == 'E')
            chinese_candidates = sum(1 for row in rows if row[6] == 'C')
            elective_counts = {'AB': 0, 'AC': 0, 'BC': 0}
            for row in rows:
                elective_counts[row[5]] += 1

            plt.figure(figsize=(15, 15))
            plt.subplot(2, 2, 1)
            plt.pie([male_candidates, female_candidates], labels=['Male', 'Female'], autopct='%1.1f%%')
            plt.title('Gender Distribution')
            plt.subplot(2, 2, 2)
            plt.pie([english_candidates, chinese_candidates], labels=['English', 'Chinese'], autopct='%1.1f%%')
            plt.title('Language Distribution')
            plt.subplot(2, 2, 3)
            plt.bar(elective_counts.keys(), elective_counts.values())
            plt.title('Elective Distribution')
            plt.ylabel('Number of Candidates')
            plt.subplot(2, 2, 4)
            plt.axis('off')
            plt.text(0.1, 0.9, f"Total Candidates: {total_candidates}", fontsize=12)
            plt.text(0.1, 0.8, f"Male Candidates: {male_candidates}", fontsize=12)
            plt.text(0.1, 0.7, f"Female Candidates: {female_candidates}", fontsize=12)
            plt.text(0.1, 0.6, f"English Candidates: {english_candidates}", fontsize=12)
            plt.text(0.1, 0.5, f"Chinese Candidates: {chinese_candidates}", fontsize=12)
            plt.tight_layout()

            plt.savefig(os.path.join(result_folder, 'statistics.png'))
            plt.close()

            with PdfPages(os.path.join(result_folder, 'statistics.pdf')) as pdf:
                fig = plt.figure(figsize=(15, 15))
                img = plt.imread(os.path.join(result_folder, 'statistics.png'))
                plt.imshow(img)
                plt.axis('off')
                pdf.savefig(fig)
                plt.close(fig)

            os.remove(os.path.join(result_folder, 'statistics.png'))

            for row in rows:
                computed_score = compute_score(row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16])
                level = determine_level(computed_score)
                cursor.execute("""
                UPDATE student_computed_grade
                SET ComputedSCORE = ?, Level = ?
                WHERE ID = ?
                """, (computed_score, level, row[0]))
            conn.commit()

            cursor.execute("SELECT DISTINCT School_name FROM student_computed_grade")
            schools = cursor.fetchall()
            for school in schools:
                school_name = school[0]
                cursor.execute("""
                SELECT Classnum, Name, Gender, Elective, Lang, Mc, Bq1, Bq2, Bq3, Bq4, Bq5, Eq1, Eq2, Eq3, Eq4, ComputedSCORE, Level
                FROM student_computed_grade
                WHERE School_name = ?
                """, (school_name,))
                school_data = cursor.fetchall()
               
                excel_file_path = os.path.join(result_folder, f"{school_name}_results.xlsx")
                if os.path.exists(excel_file_path):
                    os.remove(excel_file_path)
                wb = openpyxl.Workbook()      
                ws = wb.active
                ws.append(["Classnum", "Name", "Gender", "Elective", "Language", "MC", "BQ1", "BQ2", "BQ3", "BQ4", "BQ5", "EQ1", "EQ2", "EQ3", "EQ4", "Computed Score", "Level"])
                for student in school_data:
                    ws.append(student)
                wb.save(excel_file_path)
               
                csv_file_path = os.path.join(result_folder, f"{school_name}_results.csv")
                with open(csv_file_path, 'w', newline='') as csvfile:
                    csvwriter = csv.writer(csvfile)
                    csvwriter.writerow(["Classnum", "Name", "Gender", "Elective", "Language", "MC", "BQ1", "BQ2", "BQ3", "BQ4", "BQ5", "EQ1", "EQ2", "EQ3", "EQ4", "Computed Score", "Level"])
                    csvwriter.writerows(school_data)

            conn.close()

            self.master.after(0, lambda: self.status_label.config(text="Processing completed successfully."))
        except Exception as e:
            self.master.after(0, lambda: self.status_label.config(text=f"Error: {str(e)}"))


if __name__ == "__main__":
    root = tk.Tk()
    app = MarkSubmissionSystem(root)
    root.mainloop()