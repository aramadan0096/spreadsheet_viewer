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
        if not Path(file_path).exists():
            return []
        with open(file_path, newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            return [row for row in reader]

    def save_csv(self, file_path, data):
        """
        Save a list of lists (rows) to a CSV file.
        """
        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            for row in data:
                writer.writerow(row)
