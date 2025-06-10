import toml
from pathlib import Path

class StylingParser:
    """
    Parses Styling.toml and extracts style node lists for use in dropdowns and info panels.
    """
    def __init__(self):
        pass

    def parse_styling_file(self, file_path):
        """
        Parse the TOML styling file and return a dict with lists of style names.
        Returns:
            {
                'page_styles': [...],
                'content_styles': [...],
                'letter_styles': [...],
            }
        """
        if not Path(file_path).exists():
            return {'page_styles': [], 'content_styles': [], 'letter_styles': []}
        data = toml.load(file_path)
        # Page styles
        page_styles = [s['name'] for s in data.get('pageStyle', []) if 'name' in s]
        # Content styles
        content_styles = [s['name'] for s in data.get('contentStyle', []) if 'name' in s]
        # Letter styles
        letter_styles = [s['name'] for s in data.get('letterStyle', []) if 'name' in s]
        return {
            'page_styles': page_styles,
            'content_styles': content_styles,
            'letter_styles': letter_styles,
        }
