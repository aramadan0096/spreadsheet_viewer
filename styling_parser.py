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
        Parse the TOML styling file and return a dict with lists of style names and other values.
        """
        if not Path(file_path).exists():
            return {
                'page_styles': [],
                'content_styles': [],
                'letter_styles': [],
                'harmonization_values': [],
                'spine_positions': [],
                'gaps': [],
                'runtimes': []
            }
            
        data = toml.load(file_path)
        
        # Extract styles
        page_styles = [s['name'] for s in data.get('pageStyle', []) if 'name' in s]
        content_styles = [s['name'] for s in data.get('contentStyle', []) if 'name' in s]
        letter_styles = [s['name'] for s in data.get('letterStyle', []) if 'name' in s]
        
        # Extract harmonization values from content styles
        harmonization_values = set()
        for style in data.get('contentStyle', []):
            # Look for any *harmonize* fields that have values
            for key, value in style.items():
                if 'harmonize' in key.lower() and isinstance(value, str):
                    harmonization_values.add(value)
        harmonization_values = sorted(list(harmonization_values))
        if 'OFF' not in harmonization_values:
            harmonization_values.insert(0, 'OFF')
            
        # Extract spine positions
        spine_positions = ['BODY_CENTER', 'HEAD_GAP_CENTER', 'BODY_LEFT', 'BODY_RIGHT']
        
        # Extract gaps from global settings
        unit_gap = data.get('global', {}).get('unitVGapPx', 32.0)
        gaps = ['0', str(int(unit_gap/2)), str(int(unit_gap)), 
                str(int(unit_gap*1.5)), str(int(unit_gap*2))]
        
        # Extract runtimes from pageStyle
        runtimes = set()
        for style in data.get('pageStyle', []):
            if style.get('behavior') == 'CARD' and 'cardRuntimeFrames' in style:
                runtimes.add(str(style['cardRuntimeFrames']))
        runtimes = sorted(list(runtimes))
        if '24' not in runtimes:  # Add some standard values
            runtimes.extend(['24', '48', '72', '96'])
        runtimes = sorted(list(set(runtimes)), key=lambda x: int(x))
        
        return {
            'page_styles': page_styles,
            'content_styles': content_styles,
            'letter_styles': letter_styles,
            'harmonization_values': harmonization_values,
            'spine_positions': spine_positions,
            'gaps': gaps,
            'runtimes': runtimes
        }
