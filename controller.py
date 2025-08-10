"""
Controller for CredGen Spreadsheet Editor.
Coordinates between UI widgets and data logic.
"""
from typing import Optional
from file_manager import FileManager
from styling_parser import StylingParser

class CredGenController:
    """
    Mediates between the main window, spreadsheet widget, and data managers.
    Handles project state, undo/redo, and validation.
    """
    def __init__(self):
        self.file_manager = FileManager()
        self.styling_parser = StylingParser()
        self.current_csv_file: Optional[str] = None
        self.current_styling_file: Optional[str] = None
        self.styling_data: Optional[dict] = None
        self.undo_stack = []
        self.redo_stack = []

    def load_project(self, csv_path: str, styling_path: Optional[str] = None) -> tuple:
        """
        Load a project and return (csv_data, styling_data).
        Raises ValueError on validation errors.
        """
        csv_data = self.file_manager.load_csv(csv_path)
        if not csv_data:
            raise ValueError("CSV file is empty or invalid.")
        self.current_csv_file = csv_path
        styling_data = None
        if styling_path:
            styling_data = self.styling_parser.parse_styling_file(styling_path)
            if not styling_data:
                raise ValueError("Styling TOML is invalid.")
            self.current_styling_file = styling_path
        self.styling_data = styling_data
        return csv_data, styling_data

    def save_project(self, csv_path: str, data: list) -> None:
        """
        Save the current project data to CSV.
        """
        self.file_manager.save_csv(csv_path, data)
        self.current_csv_file = csv_path

    def validate_csv(self, csv_data: list) -> bool:
        """
        Validate CSV data structure.
        """
        if not csv_data or not isinstance(csv_data, list):
            return False
        # Add more validation as needed
        return True

    def push_undo(self, state: list) -> None:
        """Push a state to the undo stack."""
        # Store a shallow copy to avoid accidental external mutation
        self.undo_stack.append([row[:] for row in state])
        # Any new action invalidates the redo stack
        self.redo_stack.clear()

    def undo(self) -> Optional[list]:
        """
        Undo last action and return previous state.
        """
        # Need at least two states to go back: current and previous
        if len(self.undo_stack) >= 2:
            current = self.undo_stack.pop()
            self.redo_stack.append(current)
            previous = self.undo_stack[-1]
            return [row[:] for row in previous]
        return None

    def redo(self) -> Optional[list]:
        """
        Redo last undone action and return state.
        """
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.undo_stack.append([row[:] for row in state])
            return [row[:] for row in state]
        return None

    # Extensibility: plugin/config pattern for new style types
    def register_style_plugin(self, plugin_func):
        """
        Register a plugin for new style types/features.
        """
        # Example: self.plugins.append(plugin_func)
        pass
