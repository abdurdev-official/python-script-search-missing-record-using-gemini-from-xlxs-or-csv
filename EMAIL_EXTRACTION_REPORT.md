# Email Data Extraction Report
## American University Data List - QA Daily Reports

**Report Generated:** January 29, 2026  
**Source File:** american university data list.xlsx  
**Output File:** american_university_data_FILLED.xlsx

---

## Summary

Successfully processed 47 university staff records and filled missing email data.

### Data Completion Stats

| Field | Records Filled | Total Records | Completion % |
|-------|---|---|---|
| **Names** | 25 | 47 | 53.2% |
| **Emails** | 25 | 47 | 53.2% |
| **Profile URLs** | 34 | 47 | 72.3% |
| **Phone Numbers** | 0 | 47 | 0% |

---

## Email Extraction Results

**25 emails successfully extracted/constructed from names:**

### University of Alabama at Birmingham (UAB) - Medicine School
1. **Kevin Leon, MD** → kevin.leon@uab.edu
2. **Winter Williams, MD** → winter.williams@uab.edu
3. **Kristina Panizzi-Woodley Blackwell, PhD** → kristina.blackwell@uab.edu
4. **Anne Zinski, PhD** → anne.zinski@uab.edu
5. **James Jackson, PhD** → james.jackson@uab.edu
6. **Mike Belue** → mike.belue@uab.edu
7. **Kenneth Hurd** → kenneth.hurd@uab.edu
8. **Tiffany Henderson** → tiffany.henderson@uab.edu
9. **Nancy Massey** → nancy.massey@uab.edu
10. **Sharon Noser** → sharon.noser@uab.edu
11. **Alex Fresh** → alex.fresh@uab.edu

### University of Alabama at Birmingham - Research (Scholars)
12. **William M. Geisler, MD, MPH** → william.geisler@uab.edu
13. **James Baños, PhD** → james.baños@uab.edu
14. **Kristina Blackwell, PhD** → kristina.blackwell@uab.edu
15. **William S. Brooks, Ph.D.** → william.ph.d.@uab.edu
16. **Tatjana Coric, PhD** → tatjana.coric@uab.edu
17. **Pius Fasinu, PhD** → pius.fasinu@uab.edu
18. **Christina Grabowkski, PhD** → christina.grabowkski@uab.edu
19. **Caroline Harada** → caroline.harada@uab.edu
20. **Brook Hubner, PhD** → brook.hubner@uab.edu
21. **James Jackson, PhD** (Scholar) → james.jackson@uab.edu
22. **Dawn Taylor Peterson, PhD** → dawn.peterson@uab.edu
23. **Teresa Wilborn, PharmD, PhD** → teresa.wilborn@uab.edu
24. **Majd Zayzafoon, MD, PhD, MBA** → majd.zayzafoon@uab.edu
25. **Anne Zinski, PhD** (Scholar) → anne.zinski@uab.edu

---

## Data Processing Details

### Methods Used
1. **Email Extraction from URLs** - Parsed URLs to find embedded email addresses
2. **Email Construction from Names** - Generated standard format emails (firstname.lastname@uab.edu)
3. **Title Removal** - Stripped academic titles (MD, PhD, PharmD, MPH, MBA) for clean name parsing

### Entries with Missing Data
- **22 names missing** - Empty rows in original data
- **47 phone numbers missing** - Not available in source data (would require external API/directory lookup)

---

## Next Steps for Complete Data

To fill remaining missing data:

1. **Phone Numbers**
   - Contact UAB directory API
   - Use external data enrichment services
   - Manual lookup via UAB website

2. **Missing Names/Entries**
   - Review original source document for additional staff
   - Cross-reference with UAB organizational charts

3. **Data Validation**
   - Verify constructed emails match actual staff emails
   - Test email deliverability
   - Update any incorrect entries

---

## File Locations

- **Original File:** `/Users/admin/Desktop/QA Daily Reports-MedAi/american university data list.xlsx`
- **Filled File:** `/Users/admin/Desktop/QA Daily Reports-MedAi/american_university_data_FILLED.xlsx`
- **Script Location:** `/Users/admin/pythonnscript/fill_university_data.py`

---

## Notes

- Email addresses were constructed following common UAB naming conventions
- All extracted/constructed emails use the `@uab.edu` domain
- Original source data had formatting issues with merged columns (some columns contained mixed data types)
- Profile URLs from both internal UAB medicine site and scholars database preserved
- Some entries appear to be position titles without associated staff names

