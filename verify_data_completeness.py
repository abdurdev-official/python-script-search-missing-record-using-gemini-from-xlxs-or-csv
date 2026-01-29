"""
Comprehensive Data Verification and Completion
"""

import pandas as pd
import re
from typing import Tuple

# Load the filled physician data
df = pd.read_excel('physician_data_january21_3pm_FILLED.xlsx')

print("=" * 100)
print("COMPREHENSIVE DATA VERIFICATION REPORT")
print("=" * 100)
print(f"\nTotal Records: {len(df):,}")

# Field analysis
print("\n" + "=" * 100)
print("FIELD-BY-FIELD ANALYSIS:")
print("=" * 100)

field_stats = {}
for col in df.columns:
    empty_count = df[col].isna().sum() + (df[col] == '').sum()
    filled_count = len(df) - empty_count
    pct_empty = (empty_count / len(df)) * 100
    pct_filled = (filled_count / len(df)) * 100
    
    field_stats[col] = {
        'filled': filled_count,
        'empty': empty_count,
        'pct_filled': pct_filled,
        'pct_empty': pct_empty
    }
    
    print(f"\n{col.upper()}")
    print(f"  Filled: {filled_count:,} ({pct_filled:6.2f}%)")
    print(f"  Empty:  {empty_count:,} ({pct_empty:6.2f}%)")

# Data quality checks
print("\n" + "=" * 100)
print("DATA QUALITY CHECKS:")
print("=" * 100)

print(f"\nDuplicate Names: {df['name'].duplicated().sum()}")

valid_emails = df['email'].notna() & df['email'].str.contains('@', na=False)
print(f"Valid Email Format: {valid_emails.sum():,} / {df['email'].notna().sum():,}")

valid_links = df['links'].notna() & df['links'].str.contains('http', na=False)
print(f"Valid Links (URLs): {valid_links.sum():,} / {df['links'].notna().sum():,}")

# Missing critical fields
print("\n" + "=" * 100)
print("MISSING CRITICAL FIELDS:")
print("=" * 100)

missing_name = (df['name'].isna() | (df['name'] == '')).sum()
missing_email = (df['email'].isna() | (df['email'] == '')).sum()
missing_dept = (df['department'].isna() | (df['department'] == '')).sum()
missing_interest = (df['research_interest'].isna() | (df['research_interest'] == '')).sum()

print(f"\nMissing Name: {missing_name}")
print(f"Missing Email: {missing_email}")
print(f"Missing Department: {missing_dept}")
print(f"Missing Research Interest: {missing_interest}")

# Show records with missing critical data
if missing_name > 0:
    print(f"\n⚠️  Records missing NAME ({missing_name}):")
    missing_name_df = df[df['name'].isna() | (df['name'] == '')][['name', 'email', 'department']].head(10)
    print(missing_name_df.to_string())

if missing_dept > 0:
    print(f"\n⚠️  Records missing DEPARTMENT ({missing_dept}):")
    missing_dept_df = df[df['department'].isna() | (df['department'] == '')][['name', 'email', 'links', 'department']].head(10)
    print(missing_dept_df.to_string())

if missing_interest > 0:
    print(f"\n⚠️  Records missing RESEARCH INTEREST ({missing_interest}):")
    missing_int_df = df[df['research_interest'].isna() | (df['research_interest'] == '')][['name', 'email', 'department', 'research_interest']].head(10)
    print(missing_int_df.to_string())

# Data completeness summary
print("\n" + "=" * 100)
print("DATA COMPLETENESS SUMMARY:")
print("=" * 100)

completely_filled = (
    ((df['name'].notna()) & (df['name'] != '')) &
    ((df['department'].notna()) & (df['department'] != '')) &
    ((df['research_interest'].notna()) & (df['research_interest'] != ''))
).sum()

print(f"\nRecords with Name + Department + Research Interest:")
print(f"  Complete: {completely_filled:,} / {len(df):,} ({(completely_filled/len(df)*100):.2f}%)")
print(f"  Incomplete: {len(df) - completely_filled:,}")

# Overall data quality score
print("\n" + "=" * 100)
print("OVERALL DATA QUALITY SCORE:")
print("=" * 100)

quality_score = 0
critical_fields = ['name', 'department', 'research_interest']
for field in critical_fields:
    filled = field_stats[field]['filled']
    quality_score += (filled / len(df)) * 100

avg_quality = quality_score / len(critical_fields)
print(f"\nAverage Completeness: {avg_quality:.2f}%")

if avg_quality >= 99:
    print("Status: ✅ EXCELLENT - Data is 99%+ complete")
elif avg_quality >= 95:
    print("Status: ✅ VERY GOOD - Data is 95%+ complete")
elif avg_quality >= 90:
    print("Status: ⚠️  GOOD - Data is 90%+ complete (some gaps remain)")
else:
    print("Status: ❌ NEEDS ATTENTION - Data quality below 90%")

# Save verification report
with open('verification_report.txt', 'w') as f:
    f.write("DATA VERIFICATION REPORT\n")
    f.write("=" * 100 + "\n")
    f.write(f"Total Records: {len(df):,}\n")
    f.write(f"Completely Filled Records: {completely_filled:,} ({(completely_filled/len(df)*100):.2f}%)\n")
    f.write(f"Average Quality Score: {avg_quality:.2f}%\n")
    f.write(f"\nCritical Field Status:\n")
    f.write(f"  Name: {field_stats['name']['filled']:,} / {len(df):,} ({field_stats['name']['pct_filled']:.2f}%)\n")
    f.write(f"  Department: {field_stats['department']['filled']:,} / {len(df):,} ({field_stats['department']['pct_filled']:.2f}%)\n")
    f.write(f"  Research Interest: {field_stats['research_interest']['filled']:,} / {len(df):,} ({field_stats['research_interest']['pct_filled']:.2f}%)\n")

print("\n✓ Verification report saved to: verification_report.txt")

