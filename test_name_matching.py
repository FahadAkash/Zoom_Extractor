#!/usr/bin/env python3
"""
Comprehensive test script for the Zoom Attendance System name matching functionality.
Tests various name variations and edge cases to verify matching accuracy.
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'ZoomExtractor'))

from ZoomExtractor.matcher import RollMatcher

def create_test_database():
    """Create a test database with sample student data."""
    # Sample database based on the Google Sheet data
    test_data = {
        "Jahid": "1",
        "Emon": "3", 
        "Fahad Akash": "8",
        "Mehedi": "10",
        "Shihab": "11",
        "Raihan": "12",
        "Moumi": "13",
        "jannatul ferdous": "14",
        "Jim": "16",
        "Sahil": "18",
        "Ulfat": "19",
        "MOMIN": "21",
        "Faria": "23",
        "Rasel": "25",
        "Mitu": "26",
        "Alvi ( roll-29)": "29",
        "Soumik Hasan": "34",
        "Joy": "35",
        "Shahnewaz": "37",
        "Rukaiya": "38",
        "Ikhtear": "39",
        "Abdul Alim": "40",
        "Umme Hani Bithe": "41",
        "Mukti": "43",
        "Nahid": "44",
        "Soaus": "45",
        "Tomal": "46",
        "Arifa": "49",
        "Samiul Yeam": "50",
        "Pranto pal": "51"
    }
    return test_data

def run_comprehensive_tests():
    """Run comprehensive tests on name matching functionality."""
    print("=" * 80)
    print("COMPREHENSIVE NAME MATCHING TEST SUITE")
    print("=" * 80)
    
    # Initialize the matcher
    matcher = RollMatcher()
    matcher.database = create_test_database()
    
    # Test cases with expected matches
    test_cases = [
        # Exact matches (have ~95% confidence due to case-insensitive matching)
        ("Fahad Akash", "Fahad Akash", "8", 95.0),
        ("Jahid", "Jahid", "1", 95.0),
        ("Emon", "Emon", "3", 95.0),
        
        # Case variations (have ~95% confidence due to case-insensitive matching)
        ("fahad akash", "Fahad Akash", "8", 95.0),
        ("FAHAD AKASH", "Fahad Akash", "8", 95.0),
        ("Fahad akash", "Fahad Akash", "8", 95.0),
        ("jahid", "Jahid", "1", 95.0),
        ("EMON", "Emon", "3", 95.0),
        
        # Spacing variations (slightly reduced confidence due to preprocessing)
        ("Fahad  Akash", "Fahad Akash", "8", 95.0),  # Double space
        ("  Fahad Akash  ", "Fahad Akash", "8", 95.0),  # Leading/trailing spaces
        
        # Minor typos (should match with high confidence)
        ("Fahad Akashh", "Fahad Akash", "8", 90.0),  # Extra letter
        ("Fahad Akas", "Fahad Akash", "8", 90.0),     # Missing letter
        ("Fahad Akssh", "Fahad Akash", "8", 90.0),    # Substitution
        
        # Partial names (do not match - below threshold)
        ("Akash", None, None, 0.0),          # Last name only
        ("Fahad", None, None, 0.0),          # First name only
        
        # Names with special characters/punctuation (slightly reduced confidence)
        ("Alvi ( roll-29)", "Alvi ( roll-29)", "29", 95.0),
        
        # Completely different names (should not match or have very low confidence)
        ("John Smith", None, None, 0.0),
        ("Random Name", None, None, 0.0),
        
        # Edge cases
        ("jannatul ferdous", "jannatul ferdous", "14", 95.0),  # Lowercase name
        ("JANNATUL FERDOUS", "jannatul ferdous", "14", 95.0),  # Uppercase version
        ("MOMIN", "MOMIN", "21", 95.0),                       # All caps name
        ("momin", "MOMIN", "21", 95.0),                       # Lowercase version
        
        # Similar names (testing discrimination)
        ("Umme Hani Bithe", "Umme Hani Bithe", "41", 80.0),
        ("Umme-Hani-Bithe", None, None, 0.0),    # Different punctuation
    ]
    
    passed = 0
    failed = 0
    
    print(f"\nRunning {len(test_cases)} test cases...\n")
    print("-" * 80)
    print(f"{'Input Name':<25} {'Expected Match':<20} {'Exp. Roll':<10} {'Min Conf.':<10} {'Actual Result':<20} {'Status'}")
    print("-" * 80)
    
    for input_name, expected_match, expected_roll, min_confidence in test_cases:
        # Perform the match
        result = matcher.match_name(input_name)
        
        # Extract actual results
        actual_match = result.get('matched_name') if result.get('status') == 'matched' else None
        actual_roll = result.get('roll') if result.get('status') == 'matched' else None
        actual_confidence = result.get('confidence', 0.0)
        
        # Check if test passes
        name_matches = (actual_match == expected_match) or (expected_match is None and actual_match is None)
        roll_matches = (actual_roll == expected_roll) or (expected_roll is None and actual_roll is None)
        confidence_ok = actual_confidence >= min_confidence
        
        test_passed = name_matches and roll_matches and confidence_ok
        
        if test_passed:
            passed += 1
            status = "PASS"
        else:
            failed += 1
            status = "FAIL"
        
        # Print result
        expected_display = expected_match if expected_match else "None"
        actual_display = actual_match if actual_match else "None"
        print(f"{input_name:<25} {expected_display:<20} {expected_roll or 'None':<10} {min_confidence:<10.1f} "
              f"{actual_display:<20} {status}")
        
        # Print details for failed tests
        if not test_passed:
            print(f"  -> Expected: match='{expected_match}', roll='{expected_roll}', min_conf={min_confidence}%")
            print(f"  -> Actual:   match='{actual_match}', roll='{actual_roll}', conf={actual_confidence:.1f}%")
    
    print("-" * 80)
    print(f"\nTest Results: {passed} passed, {failed} failed")
    print(f"Success Rate: {passed/(passed+failed)*100:.1f}%")
    
    return passed, failed

def test_edge_cases():
    """Test edge cases and special scenarios."""
    print("\n" + "=" * 80)
    print("EDGE CASE TESTING")
    print("=" * 80)
    
    matcher = RollMatcher()
    matcher.database = create_test_database()
    
    # Test empty and invalid inputs
    edge_cases = [
        ("", "Empty string"),
        ("   ", "Whitespace only"),
        (None, "None value"),
        ("123456", "Numbers only"),
        ("!@#$%", "Special characters only"),
    ]
    
    print(f"\nTesting {len(edge_cases)} edge cases...\n")
    print("-" * 60)
    print(f"{'Input':<20} {'Description':<25} {'Result'}")
    print("-" * 60)
    
    for input_value, description in edge_cases:
        try:
            if input_value is None:
                result = matcher.match_name("")
            else:
                result = matcher.match_name(input_value)
            
            status = result.get('status', 'unknown')
            matched_name = result.get('matched_name', 'None')
            print(f"{str(input_value):<20} {description:<25} {status} ({matched_name})")
        except Exception as e:
            print(f"{str(input_value):<20} {description:<25} ERROR ({str(e)})")

if __name__ == "__main__":
    print("Zoom Attendance System - Name Matching Test Suite")
    print("This script tests the accuracy of the name detection and matching system.")
    
    # Run comprehensive tests
    passed, failed = run_comprehensive_tests()
    
    # Run edge case tests
    test_edge_cases()
    
    print("\n" + "=" * 80)
    print("TEST SUMMARY")
    print("=" * 80)
    print(f"Comprehensive Tests: {passed} passed, {failed} failed")
    print(f"Overall Success Rate: {passed/(passed+failed)*100:.1f}%" if (passed+failed) > 0 else "No tests run")
    print("\nNote: You can modify this script to test with your own data and scenarios.")