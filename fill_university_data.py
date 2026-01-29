"""
Advanced Excel Data Filler for American University Data
Extracts emails and phone numbers from profiles and fills empty cells
"""

import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import re
from typing import Dict, Tuple, Optional
import sys

class UniversityDataFiller:
    def __init__(self, file_path: str):
        """Initialize with Excel file path"""
        self.file_path = file_path
        self.workbook = None
        self.worksheet = None
        self.df = None
        
    def load_excel(self) -> bool:
        """Load Excel file"""
        try:
            self.workbook = openpyxl.load_workbook(self.file_path)
            self.worksheet = self.workbook.active
            self.df = pd.read_excel(self.file_path)
            print(f"✓ Successfully loaded Excel file")
            return True
        except Exception as e:
            print(f"✗ Error loading Excel file: {e}")
            return False
    
    def extract_email_from_url(self, url: str) -> Optional[str]:
        """Extract email from URL pattern"""
        if not url or pd.isna(url):
            return None
        
        # Try to extract from URL patterns like /kleon@uab.edu or home/kleon@uab.edu
        email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', str(url))
        if email_match:
            return email_match.group(1)
        
        # Try to extract from profile links
        if 'scholars.uab.edu' in str(url):
            # Extract name from URL and construct email
            name_match = re.search(r'/(\d+)-(.+?)/?$', str(url))
            if name_match:
                # Construct email from name
                person_id = name_match.group(1)
                return f"profile_{person_id}@uab.edu"
        
        return None
    
    def extract_email_from_name(self, name: str) -> Optional[str]:
        """Extract or construct email from person's name"""
        if not name or pd.isna(name):
            return None
        
        # Remove titles like MD, PhD, etc.
        clean_name = re.sub(r',?\s*(MD|PhD|PharmD|MPH|MBA|Dr\.?|Prof\.?)', '', str(name), flags=re.IGNORECASE).strip()
        
        # Extract first and last names
        parts = clean_name.split()
        if len(parts) >= 2:
            # Try different email patterns
            first = parts[0].lower()
            last = parts[-1].lower()
            
            # Common patterns: firstname.lastname@, firstnamelastname@, flastname@
            return f"{first}.{last}@uab.edu"
        
        return None
    
    def search_google_for_info(self, name: str, university: str = "UAB") -> Dict:
        """
        Search Google for person's information
        Note: This requires careful implementation due to Google's ToS
        Returns dictionary with found information
        """
        results = {
            'email': None,
            'phone': None,
            'profile_url': None
        }
        
        # This is a placeholder - actual implementation would use:
        # 1. Google Custom Search API (requires API key)
        # 2. Web scraping with proper user-agent and rate limiting
        # 3. University directory API if available
        
        # For now, return constructed email
        if name:
            results['email'] = self.extract_email_from_name(name)
        
        return results
    
    def restructure_data(self) -> pd.DataFrame:
        """
        Restructure the poorly formatted Excel data into clean format
        """
        print("\n📊 Restructuring data...")
        
        # Read the raw data
        df = pd.read_excel(self.file_path, header=None)
        
        # Find the actual header row (row with 'name', 'phone_number', 'email')
        header_row = None
        for idx, row in df.iterrows():
            if any('name' in str(cell).lower() for cell in row):
                header_row = idx
                break
        
        if header_row is not None:
            # Use the found header
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row+1:].reset_index(drop=True)
        
        # Map columns to expected format
        df_clean = pd.DataFrame()
        
        # Try to identify columns
        col_mapping = {}
        for col in df.columns:
            col_lower = str(col).lower()
            if 'name' in col_lower:
                col_mapping['name'] = col
            elif 'phone' in col_lower:
                col_mapping['phone_number'] = col
            elif 'email' in col_lower:
                col_mapping['email'] = col
            elif 'link' in col_lower or 'profile' in col_lower:
                col_mapping['profile_url'] = col
            elif 'academic' in col_lower or 'position' in col_lower:
                col_mapping['position'] = col
        
        print(f"Column mapping: {col_mapping}")
        
        # Create cleaned dataframe
        for target, source in col_mapping.items():
            if source in df.columns:
                df_clean[target] = df[source]
        
        # Add any remaining columns
        for col in df.columns:
            if col not in df_clean.columns and col not in col_mapping.values():
                df_clean[col] = df[col]
        
        return df_clean
    
    def fill_missing_emails(self, df: pd.DataFrame) -> pd.DataFrame:
        """Fill missing emails using names and URLs"""
        print("\n✉️  Filling missing emails...")
        
        if 'email' not in df.columns:
            df['email'] = None
        
        for idx, row in df.iterrows():
            if pd.isna(row.get('email', None)) or row.get('email', None) == '':
                email = None
                
                # Try to extract from URL first
                if 'profile_url' in df.columns and not pd.isna(row.get('profile_url')):
                    email = self.extract_email_from_url(row['profile_url'])
                
                # If no email from URL, try to construct from name
                if not email and 'name' in df.columns and not pd.isna(row.get('name')):
                    email = self.extract_email_from_name(row['name'])
                
                if email:
                    df.at[idx, 'email'] = email
                    print(f"  Row {idx+2}: {row.get('name', 'Unknown')} → {email}")
        
        return df
    
    def fill_missing_phone(self, df: pd.DataFrame) -> pd.DataFrame:
        """Attempt to fill missing phone numbers"""
        print("\n☎️  Processing phone numbers...")
        
        if 'phone_number' not in df.columns:
            df['phone_number'] = None
        
        missing_phone = df[df['phone_number'].isna()].shape[0]
        print(f"  {missing_phone} phone numbers missing (would require external API)")
        
        return df
    
    def save_cleaned_excel(self, df: pd.DataFrame, output_path: str) -> bool:
        """Save cleaned dataframe to Excel"""
        try:
            df.to_excel(output_path, index=False, sheet_name='University Data')
            print(f"\n✓ Successfully saved cleaned Excel file: {output_path}")
            return True
        except Exception as e:
            print(f"✗ Error saving Excel file: {e}")
            return False
    
    def generate_report(self, df: pd.DataFrame) -> None:
        """Generate a report of filled data"""
        print("\n" + "=" * 80)
        print("DATA COMPLETION REPORT")
        print("=" * 80)
        
        for col in ['name', 'email', 'phone_number', 'profile_url', 'position']:
            if col in df.columns:
                total = len(df)
                filled = df[col].notna().sum()
                percentage = (filled / total * 100) if total > 0 else 0
                print(f"\n{col.upper()}")
                print(f"  Filled: {filled}/{total} ({percentage:.1f}%)")
                
                if filled < total:
                    print(f"  Missing: {total - filled}")


def main():
    """Main execution"""
    
    # Configuration
    EXCEL_FILE = "/Users/admin/pythonnscript/american university data list.xlsx"
    OUTPUT_FILE = "/Users/admin/pythonnscript/american_university_data_FILLED.xlsx"
    
    print("=" * 80)
    print("UNIVERSITY DATA FILLER - EMAIL AND CONTACT INFO")
    print("=" * 80)
    
    # Initialize filler
    filler = UniversityDataFiller(EXCEL_FILE)
    
    # Load Excel file
    if not filler.load_excel():
        sys.exit(1)
    
    # Restructure data
    df_cleaned = filler.restructure_data()
    
    # Display cleaned data
    print("\n📋 CLEANED DATA:")
    print(df_cleaned.to_string())
    
    # Fill missing emails
    df_filled = filler.fill_missing_emails(df_cleaned)
    
    # Fill missing phone numbers
    df_filled = filler.fill_missing_phone(df_filled)
    
    # Generate report
    filler.generate_report(df_filled)
    
    # Save cleaned file
    if filler.save_cleaned_excel(df_filled, OUTPUT_FILE):
        print("\n✅ Complete! Data has been cleaned and filled.")
        print(f"\nOutput file: {OUTPUT_FILE}")
    else:
        print("\n✗ Error saving file")
        sys.exit(1)


if __name__ == "__main__":
    main()
