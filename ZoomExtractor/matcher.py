"""
Name-to-Roll Number Matcher
Handles fuzzy matching of detected names to roll numbers
"""

from rapidfuzz import fuzz, process
import re


class RollMatcher:
    def __init__(self, threshold=75):
        """
        Initialize matcher
        
        Args:
            threshold: Minimum similarity score (0-100) for matching
        """
        self.threshold = threshold
        self.database = {}  # {name: roll_number}
        self.matched_records = {}  # {detected_name: (matched_name, roll, confidence)}
        
    def load_from_file(self, filepath):
        """
        Load names and roll numbers from text file
        
        Expected format:
        Fahad Akash 08
        John Doe 15
        
        Returns:
            Number of records loaded
        """
        self.database = {}
        count = 0
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Try to extract name and roll number
                    # Pattern: "Name ... RollNumber" (roll at end)
                    parts = line.rsplit(maxsplit=1)
                    if len(parts) >= 2:
                        name = parts[0].strip()
                        roll = parts[1].strip()
                        self.database[name] = roll
                        count += 1
                        
        except Exception as e:
            raise Exception(f"Error reading file: {e}")
        
        print(f"Loaded {count} records from {filepath}")
        return count
    
    def load_from_text(self, text):
        """
        Load from raw text string
        
        Args:
            text: Multi-line string with name-roll pairs
        """
        self.database = {}
        count = 0
        
        for line in text.split('\n'):
            line = line.strip()
            if not line:
                continue
            
            parts = line.rsplit(maxsplit=1)
            if len(parts) >= 2:
                name = parts[0].strip()
                roll = parts[1].strip()
                self.database[name] = roll
                count += 1
        
        return count
    
    def match_name(self, detected_name):
        """
        Match a detected name to database
        
        Args:
            detected_name: Name detected by OCR
            
        Returns:
            dict: {
                'matched_name': best matching name,
                'roll': roll number,
                'confidence': match score (0-100),
                'status': 'matched' or 'unknown'
            }
        """
        if not self.database:
            return {
                'matched_name': None,
                'roll': 'N/A',
                'confidence': 0,
                'status': 'unknown'
            }
        
        # Use fuzzy matching
        result = process.extractOne(
            detected_name,
            self.database.keys(),
            scorer=fuzz.token_sort_ratio
        )
        
        if result:
            matched_name, score, _ = result
            
            if score >= self.threshold:
                return {
                    'matched_name': matched_name,
                    'roll': self.database[matched_name],
                    'confidence': score,
                    'status': 'matched'
                }
        
        return {
            'matched_name': detected_name,
            'roll': 'N/A',
            'confidence': 0,
            'status': 'unknown'
        }
    
    def match_batch(self, detected_names):
        """
        Match multiple names at once
        
        Args:
            detected_names: List of names
            
        Returns:
            dict: {detected_name: match_result}
        """
        results = {}
        for name in detected_names:
            results[name] = self.match_name(name)
            self.matched_records[name] = results[name]
        
        return results
    
    def get_all_matches(self):
        """Get all matched records"""
        return dict(self.matched_records)
    
    def get_statistics(self):
        """Get matching statistics"""
        total = len(self.matched_records)
        matched = sum(1 for r in self.matched_records.values() if r['status'] == 'matched')
        unknown = total - matched
        
        return {
            'total_detected': total,
            'matched': matched,
            'unknown': unknown,
            'match_rate': (matched / total * 100) if total > 0 else 0
        }
    
    def export_attendance(self):
        """
        Export attendance data
        
        Returns:
            List of dicts for DataFrame conversion
        """
        data = []
        for detected_name, match in self.matched_records.items():
            data.append({
                'Detected Name': detected_name,
                'Matched Name': match['matched_name'] or 'Unknown',
                'Roll Number': match['roll'],
                'Confidence': f"{match['confidence']:.1f}%",
                'Status': match['status'].title()
            })
        
        return data
