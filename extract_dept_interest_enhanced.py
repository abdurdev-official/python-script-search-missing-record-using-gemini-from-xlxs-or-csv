"""
Enhanced Department and Research Interest Extraction
"""

import pandas as pd
import re
from typing import Tuple

def clean_position_text(text: str) -> str:
    """Clean position text by removing URLs and extra whitespace"""
    if not text or pd.isna(text):
        return ""
    
    text = str(text)
    # Remove URLs
    text = re.sub(r'https?://[^\s]+', '', text)
    # Remove email addresses
    text = re.sub(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', '', text)
    # Clean up extra whitespace
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def extract_department(name: str, position_text: str) -> str:
    """Extract department from position/name info"""
    
    if not position_text or pd.isna(position_text):
        return ""
    
    pos_lower = str(position_text).lower()
    
    # Department patterns - more specific
    department_keywords = {
        'Medicine': ['medicine', 'physician', 'clinical', 'medical'],
        'Medical Education': ['medical education', 'ume', 'medical students', 'curriculum'],
        'Epidemiology': ['epidemiology', 'public health'],
        'Pathology': ['pathology'],
        'Pediatrics': ['pediatrics', 'children'],
        'Surgery': ['surgery', 'surgical'],
        'Pulmonology': ['pulmonary', 'pulmonology', 'respiratory'],
        'Preventative Medicine': ['preventative', 'preventive'],
        'Research': ['research', 'scholar'],
        'Administration': ['director', 'administrator', 'coordinator', 'manager', 'assistant'],
    }
    
    found_departments = []
    for dept, keywords in department_keywords.items():
        for keyword in keywords:
            if keyword in pos_lower:
                found_departments.append(dept)
                break
    
    # Remove duplicates while preserving order
    seen = set()
    unique_depts = []
    for dept in found_departments:
        if dept not in seen:
            unique_depts.append(dept)
            seen.add(dept)
    
    # Return most specific department
    if unique_depts:
        # Prioritize Medical Education and Epidemiology over generic Medicine
        for priority in ['Medical Education', 'Epidemiology', 'Pulmonology', 'Preventative Medicine']:
            if priority in unique_depts:
                return priority
        return unique_depts[0]
    
    return ""


def extract_research_interest(name: str, position_text: str) -> str:
    """Extract research interests"""
    
    if not position_text or pd.isna(position_text):
        return ""
    
    pos_text = str(position_text)
    pos_lower = pos_text.lower()
    
    # Keywords that indicate research interests
    interest_patterns = [
        r'research\s+([^,.]+)',
        r'scholarly\s+([^,.]+)',
        r'interest[s]?\s+(?:in|:)?\s*([^,.]+)',
        r'focus(?:es)?\s+(?:on|in)\s*([^,.]+)',
        r'(educational|clinical|medical)\s+research',
        r'(physician scientist)',
        r'(global initiatives)',
        r'(international[^,\.]*)',
        r'(pathways[^,\.]*)',
        r'(clinical skills)',
        r'(assessment[^,\.]*)',
        r'(medical education)',
    ]
    
    interests = []
    for pattern in interest_patterns:
        matches = re.findall(pattern, pos_lower)
        for match in matches:
            if isinstance(match, str):
                cleaned = match.strip()
                # Only add if it's meaningful (more than 2 words typically)
                if cleaned and cleaned not in interests:
                    interests.append(cleaned)
    
    # Combine interests
    if interests:
        # Remove duplicates
        unique = []
        seen = set()
        for i in interests:
            if i not in seen:
                unique.append(i)
                seen.add(i)
        return ", ".join(unique[:2])  # Return up to 2 interests
    
    return ""


# Load the data
print("=" * 90)
print("ENHANCED DEPARTMENT AND RESEARCH INTEREST EXTRACTION")
print("=" * 90)

df = pd.read_excel('american_university_data_FILLED.xlsx')

# Add/clear columns
df['department'] = ""
df['research_interest'] = ""

# Process each record
dept_count = 0
interest_count = 0

for idx, row in df.iterrows():
    name = row.get('name', '')
    if pd.isna(name) or name == '':
        continue
    
    # Combine position texts
    position_text = ""
    for col in ['link_to_profile', 'profile_url']:
        if col in df.columns:
            val = row.get(col, '')
            if pd.notna(val) and val != '' and 'http' not in str(val):
                position_text += " " + str(val)
    
    # Clean the text
    clean_pos = clean_position_text(position_text)
    
    # Extract department
    dept = extract_department(name, clean_pos)
    if dept:
        df.at[idx, 'department'] = dept
        dept_count += 1
    
    # Extract research interest
    interest = extract_research_interest(name, clean_pos)
    if interest:
        df.at[idx, 'research_interest'] = interest
        interest_count += 1

print(f"\n✓ Extracted {dept_count} departments")
print(f"✓ Extracted {interest_count} research interests")

# Display results
print("\n" + "=" * 90)
print("RESULTS WITH DEPARTMENT AND RESEARCH INTERESTS:")
print("=" * 90)

for idx, row in df.iterrows():
    if pd.notna(row.get('name')) and row.get('name') != '':
        print(f"\n{idx + 1:2}. {row['name']}")
        if row.get('department'):
            print(f"    📍 Department: {row['department']}")
        if row.get('research_interest'):
            print(f"    🔬 Research: {row['research_interest']}")
        if row.get('email'):
            print(f"    ✉️  Email: {row['email']}")

# Save updated file
output_file = 'american_university_data_FILLED.xlsx'
df.to_excel(output_file, index=False, sheet_name='University Data')
print(f"\n✓ Saved to: {output_file}")

# Copy to Desktop
import shutil
desktop_path = "/Users/admin/Desktop/QA Daily Reports-MedAi/american_university_data_FILLED.xlsx"
shutil.copy(output_file, desktop_path)
print(f"✓ Copied to Desktop: {desktop_path}")

# Final summary
print("\n" + "=" * 90)
print("FINAL DATA SUMMARY")
print("=" * 90)
print(f"Total Records: {len(df)}")
print(f"With Names: {df['name'].notna().sum()}")
print(f"With Emails: {df['email'].notna().sum()}")
print(f"With Departments: {(df['department'] != '').sum()}")
print(f"With Research Interests: {(df['research_interest'] != '').sum()}")
print(f"\nColumns: {', '.join(df.columns.tolist())}")
