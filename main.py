import os
from database.importer import import_and_process_file
from database.connect import initialize_database

def main(upload_folder):
    """
    Main function to initialize the database and process all files in the given upload folder.
    """
    print("Initializing the database...")
    initialize_database()  # Set up the initial database connection and schema.

    print("Database initialized successfully.\nStarting file processing...")
    
    # Initialize a counter to track the number of processed files
    processed_files_count = 0
    
    # Iterates through each file in the specified directory
    for filename in os.listdir(upload_folder):
        # Ensures processing only for Excel or CSV files
        if filename.endswith(('.xlsx', '.csv')):
            try:
                print(f"Processing file: {filename}")
                
                # Constructs the full path to the target file
                file_path = os.path.join(upload_folder, filename)
                
                # Processes the file, including reading, transforming, and saving its data
                import_and_process_file(file_path)
                
                # Indicates successful processing of the file
                print(f"Successfully processed file: {filename}")
                processed_files_count += 1

            except Exception as e:
                # Handles any errors during file processing by logging them
                print(f"An error occurred while processing {filename}: {e}")
    
    # Provides a final report once all files have been processed
    print(f"File processing complete. Total files processed: {processed_files_count}")

if __name__ == "__main__":
    # Specifies the directory containing files to be processed
    # Be sure to replace the path with your actual directory path
    upload_folder_path = "C:\\Users\\yukli\\OneDrive\\upload"
    main(upload_folder_path)