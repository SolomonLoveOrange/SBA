import sqlite3

# Define your database filename
DATABASE_FILE = 'hkdse_ict_system.db'

def connect_db():
    """Create a database connection to the SQLite database specified by DATABASE_FILE."""
    return sqlite3.connect(DATABASE_FILE)

def initialize_database():
    """Initializes the database by creating the necessary tables if they don't already exist."""
    conn = connect_db()
    cursor = conn.cursor()

    # SQL statement to create a table
    create_table_sql = """
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_id TEXT NOT NULL,
        Classnum TEXT NOT NULL,
        Name TEXT NOT NULL,
        Gender TEXT NOT NULL,
        Elective TEXT NOT NULL,
        Lang TEXT NOT NULL,
        MC INTEGER,
        Bq1 INTEGER,
        Bq2 INTEGER,
        Bq3 INTEGER,
        Bq4 INTEGER,
        Bq5 INTEGER,
        Eq1 INTEGER,
        Eq2 INTEGER,
        Eq3 INTEGER,
        Eq4 INTEGER,
        total_score FLOAT,
        level TEXT
    );
    """
    # Execute the statement to create the table
    cursor.execute(create_table_sql)

    # Save the changes and close the connection
    conn.commit()
    conn.close()

def insert_score(school_id, classnum, name, gender, elective, lang, mc, bq1, bq2, bq3, bq4, bq5, eq1, eq2, eq3, eq4):
    """Insert a new student record into the students table."""
    conn = connect_db()
    cursor = conn.cursor()
    
    cursor.execute('''
        INSERT INTO students 
        (school_id, classnum, name, gender, elective, lang, mc, bq1, bq2, bq3, bq4, bq5, eq1, eq2, eq3, eq4)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    ''', (school_id, classnum, name, gender, elective, lang, mc, bq1, bq2, bq3, bq4, bq5, eq1, eq2, eq3, eq4))
    
    conn.commit()
    conn.close()