import csv
from pathlib import Path

class FileManager:
    """
    Handles loading and saving CSV files for the spreadsheet editor.
    """
    def __init__(self):
        pass

    def load_csv(self, file_path):
        """
        Load a CSV file and return a list of lists (rows).
        """
        try:
            if not file_path:
                print("No file path provided")
                return []
                
            path = Path(file_path)
            if not path.exists():
                print(f"File not found: {file_path}")
                return []
                
            rows = []
            print(f"Reading CSV file: {file_path}")
            
            # First read the file content and remove any comment lines
            with open(file_path, 'r', newline='', encoding='utf-8-sig') as f:
                content = f.read()
                # Remove comment lines starting with //
                content_lines = [line for line in content.splitlines() if not line.strip().startswith('//')]
                content = '\n'.join(content_lines)
            
            # Now parse the cleaned content as CSV
            import io
            csv_file = io.StringIO(content)
            reader = csv.reader(csv_file, quoting=csv.QUOTE_MINIMAL)
            rows = [row for row in reader if row]  # Skip empty rows
            
            print(f"Read {len(rows)} rows from CSV")
            if rows:
                print(f"First row: {rows[0]}")
                
            # If first row isn't headers, add them
            if not rows or not rows[0][0].startswith('@'):
                headers = ["@Head", "@Body", "@Tail", "@Vertical Gap", "@Content Style", 
                          "@Break Harmonization", "@Spine Position", "@Page Style",
                          "@Page Runtime", "@Page Gap"]
                rows.insert(0, headers)
                print("Added default headers")
                
            return rows
            
        except Exception as e:
            print(f"Error loading CSV: {str(e)}")
            raise
            

    def save_csv(self, file_path, data):
        """
        Save a list of lists (rows) to a CSV file.
        """
        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            for row in data:
                writer.writerow(row)
