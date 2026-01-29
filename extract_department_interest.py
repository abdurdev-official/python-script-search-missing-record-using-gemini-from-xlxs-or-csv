"""
Extract Department and Research Interests from University Data
"""

import pandas as pd
import re
from typing import Tuple

def extract_department_and_interest(position: str, profile_url: str = None) -> Tuple[str, str]:
    """
    Extract department and research interest from position/profile text
    Returns: (department, research_interest)
    """
    department = None
    interest = None
    
    if not position or pd.isna(position):
        return department, interest
    
    pos_text = str(position).strip()
    
    # Extract Department (look for common patterns)
    department_patterns = [
        r'Medicine\s*-\s*([^,]+)',  # Medicine - Specialty
        r'(Medicine|Epidemiology|Education)',
        r'(Medical Education|UME|SOM)',
    ]
    
    for pattern in department_patterns:
        match = re.search(pattern, pos_text, re.IGNORECASE)
        if match:
            department = match.group(1) if match.groups() else match.group(0)
            break
    
    # Extract Research Interest (look for research-related keywords)
    interest_keywords = [
        'research',
        'education',
        'scholarly activity',
        'physician scientist',
        'clinical',
        'assessment',
        'development',
        'initiatives',
        'pathways',
        'medical education',
        'global initiatives',
        'international',
    ]
    
    pos_lower = pos_text.lower()
    found_interests = []
    
    for keyword in interest_keywords:
        if keyword in pos_lower:
            # Extract surrounding text
            idx = pos_lower.find(keyword)
            start = max(0, idx - 20)
            end = min(len(pos_text), idx + len(keyword) + 30)
            context = pos_text[start:end].strip()
            if context not in found_interests:
                found_interests.append(context)
    
    if found_interests:
        # Take the most specific one (longest)
        interest = found_interests[-1]
    
    # Clean up interest text
    if interest:
        # Remove common prefixes/suffixes
        interest = re.sub(r'^[^a-zA-Z]*', '', interest)
        interest = re.sub(r'[^a-zA-Z0-9&\s,\-\.]*$', '', interest)
    
    return department, interest


# Load the filled Excel file
print("=" * 80)
print("EXTRACTING DEPARTMENT AND RESEARCH INTERESTS")
print("=" * 80)

df = pd.read_excel('american_university_data_FILLED.xlsx')

# Add new columns
df['department'] = None
df['research_interest'] = None

# Extract department and interest for each person
filled_count = 0
for idx, row in df.iterrows():
    if pd.notna(row.get('name')) and row.get('name') != '':
        # Get position text (could be in 'profile_url' or 'link_to_profile')
        position_text = row.get('link_to_profile')
        if pd.isna(position_text):
            position_text = row.get('profile_url')
        
        if position_text and not 'http' in str(position_text):
            dept, interest = extract_department_and_interest(position_text, row.get('profile_url'))
            
            if dept:
                df.at[idx, 'department'] = dept
                filled_count += 1
            
            if interest:
                df.at[idx, 'research_interest'] = interest

print(f"\n✓ Extracted {filled_count} departments/interests")

# Display the results
print("\n" + "=" * 80)
print("EXTRACTED DATA:")
print("=" * 80)

for idx, row in df.iterrows():
    if pd.notna(row.get('name')) and row.get('name') != '':
        print(f"\n{idx + 1}. {row.get('name')}")
        print(f"   Department: {row.get('department') or 'N/A'}")
        print(f"   Research Interest: {row.get('research_interest') or 'N/A'}")

# Save the updated file
output_file = 'american_university_data_FILLED.xlsx'
df.to_excel(output_file, index=False, sheet_name='University Data')
print(f"\n✓ Saved updated file: {output_file}")

# Display summary
print("\n" + "=" * 80)
print("DATA SUMMARY")
print("=" * 80)
print(f"Total records: {len(df)}")
print(f"Records with department: {df['department'].notna().sum()}")
print(f"Records with research interest: {df['research_interest'].notna().sum()}")
print(f"\nColumns in file: {list(df.columns)}")
