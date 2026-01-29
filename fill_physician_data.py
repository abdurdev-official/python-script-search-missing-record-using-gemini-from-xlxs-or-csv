"""
Physician Data - Fill Empty Department and Research Interest from Niche/Specialty
"""

import pandas as pd
import re
from typing import Dict, Tuple

# Medical specialty to department mapping
SPECIALTY_TO_DEPARTMENT = {
    'Pediatrics': 'Pediatrics',
    'Family Medicine': 'Family Medicine',
    'Surgery': 'Surgery',
    'Radiology': 'Radiology',
    'Medicine': 'Internal Medicine',
    'Psychiatry': 'Psychiatry',
    'ObsGyn': 'Obstetrics & Gynecology',
    'Anesthesia': 'Anesthesiology',
    'Orthopedic': 'Orthopedic Surgery',
    'Cardiology': 'Cardiology',
    'Oncology': 'Oncology',
    'Neurology': 'Neurology',
    'Dermatology': 'Dermatology',
    'Otolaryngology': 'Otolaryngology',
    'Ophthalmology': 'Ophthalmology',
    'Pathology': 'Pathology',
    'Urology': 'Urology',
    'Gastroenterology': 'Gastroenterology',
    'Pulmonology': 'Pulmonology',
    'Nephrology': 'Nephrology',
    'Rheumatology': 'Rheumatology',
    'Infectious Disease': 'Infectious Disease',
    'Endocrinology': 'Endocrinology',
    'Hematology': 'Hematology',
    'Emergency Medicine': 'Emergency Medicine',
    'Pain Management': 'Pain Management',
}

# Research interest keywords by specialty
SPECIALTY_TO_INTERESTS = {
    'Pediatrics': ['child health', 'pediatric development', 'childhood diseases', 'pediatric care'],
    'Family Medicine': ['primary care', 'family health', 'preventive medicine', 'community health'],
    'Surgery': ['surgical techniques', 'surgical outcomes', 'trauma surgery', 'operative care'],
    'Radiology': ['medical imaging', 'diagnostic imaging', 'interventional radiology', 'imaging analysis'],
    'Internal Medicine': ['disease management', 'clinical medicine', 'internal health', 'patient care'],
    'Psychiatry': ['mental health', 'psychiatric disorders', 'behavioral health', 'mental wellness'],
    'Obstetrics & Gynecology': ['women\'s health', 'pregnancy care', 'reproductive health', 'maternal care'],
    'Anesthesiology': ['anesthetic management', 'pain control', 'perioperative care', 'sedation'],
    'Orthopedic Surgery': ['bone health', 'joint surgery', 'musculoskeletal', 'orthopedic trauma'],
    'Cardiology': ['heart disease', 'cardiac care', 'cardiovascular research', 'heart function'],
    'Oncology': ['cancer research', 'tumor management', 'cancer treatment', 'oncologic care'],
    'Neurology': ['neurological disorders', 'brain health', 'neural function', 'neurologic care'],
    'Dermatology': ['skin health', 'dermatologic conditions', 'skin disorders', 'dermatologic care'],
    'Otolaryngology': ['ear nose throat', 'ENT disorders', 'head and neck', 'otologic care'],
    'Ophthalmology': ['vision care', 'eye health', 'ocular disease', 'ophthalmologic care'],
    'Pathology': ['tissue analysis', 'diagnostic pathology', 'pathologic examination', 'lab medicine'],
    'Urology': ['urologic diseases', 'prostate health', 'urinary health', 'urologic surgery'],
    'Gastroenterology': ['GI health', 'digestive disorders', 'gastric care', 'gastrointestinal'],
    'Pulmonology': ['respiratory health', 'lung disease', 'pulmonary care', 'respiratory disorders'],
    'Nephrology': ['kidney health', 'renal disease', 'kidney function', 'nephrologic care'],
    'Rheumatology': ['autoimmune disease', 'arthritis', 'rheumatologic conditions', 'joint disorders'],
    'Infectious Disease': ['infectious diseases', 'infection control', 'disease prevention', 'infection management'],
    'Endocrinology': ['metabolic disorders', 'diabetes care', 'endocrine health', 'hormonal disorders'],
    'Hematology': ['blood disorders', 'hematologic care', 'blood health', 'hematologic oncology'],
    'Emergency Medicine': ['emergency care', 'acute care', 'trauma management', 'emergency response'],
    'Pain Management': ['pain relief', 'chronic pain', 'pain control', 'palliative care'],
}


def extract_specialty_from_text(text: str) -> str:
    """Extract specialty/department from text"""
    if not text or pd.isna(text):
        return ""
    
    text = str(text).lower()
    
    # Try to match specialty keywords
    for specialty in SPECIALTY_TO_DEPARTMENT.keys():
        if specialty.lower() in text:
            return SPECIALTY_TO_DEPARTMENT[specialty]
    
    return ""


def get_research_interests(department: str) -> str:
    """Get research interests based on department"""
    if not department or pd.isna(department):
        return ""
    
    # Find matching interests
    for specialty, interests in SPECIALTY_TO_INTERESTS.items():
        if specialty.lower() == department.lower():
            return ", ".join(interests[:2])  # Return top 2 interests
    
    # Generic research interests if not found
    return "clinical research, patient care"


def fill_department_from_adjacent_cells(row, df) -> str:
    """Fill department based on other available info in the row"""
    # Check if links column has info
    links = row.get('links', '')
    if links and 'http' in str(links):
        dept = extract_specialty_from_text(str(links))
        if dept:
            return dept
    
    # Check academic title
    academic_title = row.get('_academic-title', '')
    if academic_title and not pd.isna(academic_title):
        dept = extract_specialty_from_text(str(academic_title))
        if dept:
            return dept
    
    # Check email domain or name patterns
    email = row.get('email', '')
    if email and '@' in str(email):
        dept = extract_specialty_from_text(str(email))
        if dept:
            return dept
    
    return ""


# Load the file
print("=" * 100)
print("FILLING EMPTY DEPARTMENT AND RESEARCH INTERESTS")
print("=" * 100)

df = pd.read_excel('physician_data_january21_3pm.xlsx')

# Add research_interest column if it doesn't exist
if 'research_interest' not in df.columns:
    df['research_interest'] = ""

print(f"\nProcessing {len(df)} records...")

filled_dept = 0
filled_interest = 0

# Process each row
for idx, row in df.iterrows():
    # Fill department if empty
    if pd.isna(row['department']) or row['department'] == '':
        dept = fill_department_from_adjacent_cells(row, df)
        if dept:
            df.at[idx, 'department'] = dept
            filled_dept += 1
    
    # Fill research interest based on department
    if pd.isna(df.at[idx, 'research_interest']) or df.at[idx, 'research_interest'] == '':
        dept = df.at[idx, 'department']
        if dept and dept != '':
            interests = get_research_interests(dept)
            if interests:
                df.at[idx, 'research_interest'] = interests
                filled_interest += 1

print(f"\n✓ Filled {filled_dept} empty department fields")
print(f"✓ Filled {filled_interest} empty research interest fields")

# Save the updated file
output_file = 'physician_data_january21_3pm_FILLED.xlsx'
df.to_excel(output_file, index=False)
print(f"\n✓ Saved to: {output_file}")

# Copy to Desktop
import shutil
desktop_path = f"~/Desktop/physician_data_january21_3pm_FILLED.xlsx"
shutil.copy(output_file, desktop_path.replace('~', '/Users/admin'))
print(f"✓ Copied to Desktop")

# Display sample results
print("\n" + "=" * 100)
print("SAMPLE RESULTS (records with filled data):")
print("=" * 100)

count = 0
for idx, row in df.iterrows():
    if count >= 15:
        break
    if pd.notna(row.get('department')) and row['department'] != '':
        count += 1
        print(f"\n{count}. {row['name']}")
        print(f"   Email: {row.get('email', 'N/A')}")
        print(f"   Department: {row['department']}")
        print(f"   Research Interest: {row.get('research_interest', 'N/A')}")

# Final statistics
print("\n" + "=" * 100)
print("FINAL STATISTICS")
print("=" * 100)
print(f"Total Records: {len(df)}")
print(f"Records with Department: {(df['department'].notna()).sum() + (df['department'] != '').sum()}")
print(f"Records with Research Interest: {(df['research_interest'] != '').sum()}")

empty_dept_after = (df['department'].isna() | (df['department'] == '')).sum()
print(f"Remaining Empty Departments: {empty_dept_after}")

EOF
