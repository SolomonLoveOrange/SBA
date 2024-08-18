Mark Submission System
======================

This system is designed for the Hong Kong IT Society's ICT mock exam, automating the process of submitting and calculating student scores.

Requirements
------------
- Python 3.12
- SQLite
- Libraries (install via pip):
  - openpyxl==3.1.3
  - pandas==2.2.2
  - matplotlib==3.9.1
  - tkinter (usually comes with Python)
  -os (usually comes with Python)
  -csv (usually comes with Python)
  -threading (usually comes with Python)

Installation
------------
It's recommended to use a virtual environment for this project. This isolates your project dependencies from other Python projects.

1. Ensure Python 3.12 is installed on your system.
2. Clone or download this repository.
3. Navigate to the project directory.
4. Create a virtual environment:
   python -m venv venv

5. Activate the virtual environment:
   - On Windows:
     venv\Scripts\activate
   - On macOS and Linux:
     source venv/bin/activate

6. Install required libraries:
   pip install -r requirements.txt

Usage
-----
1. Ensure your virtual environment is activated.

2. Run the main script:
   python mark_submission_system.py

3. Use the GUI to:
   - Select input folder (containing .xlsx or .csv files)
   - Select output folder (for results and statistics)
   - Click "Process Data" to start

4. The system will:
   - Import data from input files
   - Calculate scores and levels
   - Generate statistics and visualizations
   - Export results for each school

5. When finished, you can deactivate the virtual environment:
   deactivate

File Structure
--------------
- mark_submission_system.py: Main script
- requirements.txt: List of required Python libraries
- README.txt: This file
- /upload: Place input files here
- /result: Output files will be saved here

Input File Format
-----------------
- Excel (.xlsx) or CSV files
- Filename format: [school_id].xlsx or [school_id].csv
- Columns: Classnum, Name, Gender, Elective, Lang, Mc, Bq1, Bq2, Bq3, Bq4, Bq5, Eq1, Eq2, Eq3, Eq4

Output
------
- Excel and CSV files with computed scores and levels for each school
- PDF file with statistics and visualizations
- SQLite database (in the /database folder) containing all processed data

Notes
-----
- Ensure input files are in the correct format
- The system uses SQLite for data storage, with the database file created in the /result/database folder
- Cut scores for levels are predefined in the code

Troubleshooting
---------------
If you encounter any issues:
1. Check that all required libraries are installed
2. Ensure input files are in the correct format and location
3. Verify write permissions for the output folder
4. Make sure you're running the script from within the activated virtual environment

For further assistance, contact: solomon02092007@gmail.com
