"""
Name-to-Roll Number Matcher
Handles fuzzy matching of detected names to roll numbers
"""

from rapidfuzz import fuzz, process
import re
import requests
import pandas as pd


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
        
        Supports multiple formats:
        - "Fahad Akash 08" (Name followed by roll)
        - "1. Jahid" or "1	Jahid" (Roll followed by name)
        
        Returns:
            Number of records loaded
        """
        self.database = {}
        count = 0
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Try Format 1: "1. Name" or "1	Name" or "1 Name"
                    # Pattern: number (with optional dot/tab) followed by name
                    match = re.match(r'^(\d+)[\.\s\t]+(.+)$', line)
                    if match:
                        roll = match.group(1).strip()
                        name = match.group(2).strip()
                        self.database[name] = roll
                        count += 1
                        print(f"  Line {line_num}: '{name}' → Roll {roll}")
                        continue
                    
                    # Try Format 2: "Name RollNumber" (roll at end)
                    parts = line.rsplit(maxsplit=1)
                    if len(parts) >= 2:
                        # Check if last part is a number or could be roll number
                        potential_roll = parts[1].strip()
                        name = parts[0].strip()
                        
                        # Accept if it looks like a roll number
                        if potential_roll.isdigit() or len(potential_roll) <= 3:
                            self.database[name] = potential_roll
                            count += 1
                            print(f"  Line {line_num}: '{name}' → Roll {potential_roll}")
                            continue
                    
                    # If we get here, format not recognized
                    print(f"  Line {line_num}: Skipped (unrecognized format): {line}")
                        
        except Exception as e:
            raise Exception(f"Error reading file: {e}")
        
        print(f"\n✓ Loaded {count} records from {filepath}")
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
        
        Priority:
        1. Roll number found in text (e.g. "15 Fahad") -> 100% match
        2. Fuzzy name matching (e.g. "Fahad") -> Score based match
        """
        if not self.database:
            return {
                'matched_name': None,
                'roll': 'N/A',
                'confidence': 0,
                'status': 'unknown'
            }
        
        # Preprocess the detected name
        processed_name = self.preprocess_name(detected_name)
        print(f"Matching name: '{detected_name}' -> processed: '{processed_name}'")
        
        # STRATEGY 1: Look for Roll Number in the Zoom Name
        # Only do this if the detected name looks like it might contain a roll number
        # (e.g., "3. Name" or "Name 3" but not "Participants (3)")
        import re
        
        # Create a reverse lookup map (Roll -> Name)
        roll_to_name = {v: k for k, v in self.database.items()}
        
        # Look for roll numbers in common formats
        # Format 1: "3. Name" or "3 Name"
        # Format 2: "Name 3"
        
        # Check if this looks like a numbered list item (e.g., "3. Name")
        list_pattern = r'^(\d+)\.?\s+(.+)$'
        list_match = re.match(list_pattern, detected_name.strip())
        
        if list_match:
            roll_num = list_match.group(1)
            name_part = list_match.group(2)
            # Check if this roll number exists in our database
            if roll_num in roll_to_name:
                matched_db_name = roll_to_name[roll_num]
                print(f"  ✓ FOUND ROLL NUMBER MATCH (list format): {roll_num} -> {matched_db_name}")
                return {
                    'matched_name': matched_db_name,
                    'roll': self.database[matched_db_name],
                    'confidence': 100,  # Perfect match by ID
                    'status': 'matched'
                }
        
        # Check if this ends with a roll number (e.g., "Name 3")
        # But avoid matching things like "Participants (3)" or "Room 3"
        end_pattern = r'^(.+)\s+(\d+)$'
        end_match = re.match(end_pattern, detected_name.strip())
        
        if end_match:
            name_part = end_match.group(1)
            roll_num = end_match.group(2)
            # Only accept if the roll number is in our database AND the name part looks reasonable
            # (avoid matching generic terms like "Participants", "Room", "Meeting", etc.)
            forbidden_words = ['participant', 'meeting', 'room', 'group', 'section', 'level', 'session']
            name_lower = name_part.lower()
            is_forbidden = any(word in name_lower for word in forbidden_words)
            
            if roll_num in roll_to_name and len(name_part) > 3 and not is_forbidden:
                matched_db_name = roll_to_name[roll_num]
                print(f"  ✓ FOUND ROLL NUMBER MATCH (end format): {roll_num} -> {matched_db_name}")
                return {
                    'matched_name': matched_db_name,
                    'roll': self.database[matched_db_name],
                    'confidence': 100,  # Perfect match by ID
                    'status': 'matched'
                }

        # STRATEGY 2: Fuzzy Name Matching
        result = process.extractOne(
            processed_name,
            self.database.keys(),
            scorer=fuzz.token_sort_ratio
        )
        
        if result:
            matched_name, score, _ = result
            print(f"  Found fuzzy match: '{matched_name}' with score {score}")
            
            if score >= self.threshold:
                print(f"  Match accepted (threshold: {self.threshold})")
                return {
                    'matched_name': matched_name,
                    'roll': self.database[matched_name],
                    'confidence': score,
                    'status': 'matched'
                }
            else:
                print(f"  Match rejected (below threshold: {self.threshold})")
        
        # If no match found, try with original name as fallback
        if processed_name != detected_name:
            print(f"  Trying fallback with original name: '{detected_name}'")
            result = process.extractOne(
                detected_name,
                self.database.keys(),
                scorer=fuzz.token_sort_ratio
            )
            
            if result:
                matched_name, score, _ = result
                print(f"  Fallback found match: '{matched_name}' with score {score}")
                
                if score >= self.threshold:
                    print(f"  Fallback match accepted (threshold: {self.threshold})")
                    return {
                        'matched_name': matched_name,
                        'roll': self.database[matched_name],
                        'confidence': score,
                        'status': 'matched'
                    }
                else:
                    print(f"  Fallback match rejected (below threshold: {self.threshold})")
            else:
                print("  No match found with original name")
        
        print(f"  No match found, returning as unknown")
        return {
            'matched_name': detected_name,
            'roll': 'N/A',
            'confidence': 0,
            'status': 'unknown'
        }
    
    def preprocess_name(self, name):
        """Preprocess name to handle common OCR issues"""
        if not name:
            return ""
            
        # 1. Remove leading/trailing noise characters
        name = name.strip('| [](){}<>.,:;-_')
        
        # 2. Remove specific noise patterns using regex
        # Remove vertical bars inside text
        name = name.replace('|', '')
        
        # Remove (Host, me) and variations - handles missing opening parenthesis too
        name = re.sub(r'\(?.*?(?:Host|Me|me).*?\)?', '', name, flags=re.IGNORECASE)
        
        # Remove "GBR", "ED" and other common short noise words
        name = re.sub(r'\b(?:GBR|ED|ER|ft)\b', '', name, flags=re.IGNORECASE)
        
        # Clean up multiple spaces
        name = ' '.join(name.split())
        
        # 3. Apply specific name corrections
        corrections = {
            'Farhad': 'Fahad',
            'Farad': 'Fahad',
            'atash': 'Akash',
            'Akas': 'Akash',
            'Akashhh': 'Akash',
            'Akashh': 'Akash',
            'Fahad Akas': 'Fahad Akash',
        }
        
        # Apply corrections
        for wrong, correct in corrections.items():
            if wrong in name:
                name = name.replace(wrong, correct)
        
        return name.strip()
    
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
    
    def load_from_google_sheet(self, sheet_url):
        """
        Load names and roll numbers from Google Sheet
        
        Args:
            sheet_url: Google Sheets URL
        
        Returns:
            Number of records loaded
        """
        try:
            # Convert Google Sheets URL to CSV export URL
            # Handle different URL formats
            if '/edit' in sheet_url:
                # Extract the document ID
                doc_id = sheet_url.split('/d/')[1].split('/')[0]
                csv_url = f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=csv"
            elif '/export' in sheet_url:
                csv_url = sheet_url.split('&')[0]  # Remove extra parameters
            else:
                raise ValueError("Invalid Google Sheets URL format")
            
            # Load CSV data using pandas
            df = pd.read_csv(csv_url)
            
            # Clear existing database
            self.database = {}
            count = 0
            
            # Remove completely empty rows and columns
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # If dataframe is empty, raise an error
            if df.empty:
                raise ValueError("No data found in Google Sheet")
            
            # Reset index after dropping rows
            df = df.reset_index(drop=True)
            
            # Find the header row (first row with 'name' and 'roll' or similar)
            header_row_idx = 0
            name_col_idx = -1
            roll_col_idx = -1
            
            # Look for header row in first 5 rows
            for idx in range(min(5, len(df))):
                row_values = [str(val).strip().lower() for val in df.iloc[idx] if pd.notna(val)]
                row_str = ' '.join(row_values)
                
                # Check if this row contains header-like values
                if 'name' in row_str or 'roll' in row_str or 'id' in row_str:
                    # Find name and roll columns in this row
                    for col_idx, val in enumerate(df.iloc[idx]):
                        if pd.isna(val):
                            continue
                        val_str = str(val).strip().lower()
                        if 'name' in val_str:
                            name_col_idx = col_idx
                        elif 'roll' in val_str or 'id' in val_str:
                            roll_col_idx = col_idx
                    
                    if name_col_idx != -1 and roll_col_idx != -1:
                        header_row_idx = idx
                        break
            
            # If we couldn't find proper headers, use the first row as data and assume first two columns
            if name_col_idx == -1 or roll_col_idx == -1:
                name_col_idx = 0
                roll_col_idx = min(1, len(df.columns) - 1)
                header_row_idx = -1  # No header row
            
            # Slice dataframe to start from after header row
            data_start_idx = header_row_idx + 1 if header_row_idx >= 0 else 0
            df_data = df.iloc[data_start_idx:].reset_index(drop=True)
            
            # Get column names
            if header_row_idx >= 0 and name_col_idx < len(df.columns) and roll_col_idx < len(df.columns):
                actual_name_col = df.columns[name_col_idx]
                actual_roll_col = df.columns[roll_col_idx]
            else:
                # Use positional indexing
                actual_name_col = df_data.columns[name_col_idx]
                actual_roll_col = df_data.columns[roll_col_idx]
            
            # Process each row
            for _, row in df_data.iterrows():
                # Skip if either column is missing
                if name_col_idx >= len(row) or roll_col_idx >= len(row):
                    continue
                    
                name_val = row.iloc[name_col_idx] if name_col_idx < len(row) else None
                roll_val = row.iloc[roll_col_idx] if roll_col_idx < len(row) else None
                
                # Convert to string and strip
                name = str(name_val).strip() if pd.notna(name_val) else ''
                roll = str(roll_val).strip() if pd.notna(roll_val) else ''
                
                # Skip empty rows
                if name and roll and name.lower() != 'nan' and roll.lower() != 'nan' and name != 'None' and roll != 'None':
                    # Additional check to ensure we have meaningful data
                    if len(name) > 1 and len(roll) >= 1:
                        self.database[name] = roll
                        count += 1
            
            print(f"\n✓ Loaded {count} records from Google Sheet")
            print("Loaded records:")
            for name, roll in list(self.database.items())[:10]:  # Show first 10 records
                print(f"  {name} -> {roll}")
            if len(self.database) > 10:
                print(f"  ... and {len(self.database) - 10} more records")
            return count
            
        except Exception as e:
            raise Exception(f"Error loading Google Sheet: {e}")
    
