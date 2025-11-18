# Alkox Washout Matrix - Python Conversion

## Overview
This project converts the VBA macro from the Excel file `Alkox Washout Matrix Product Wheel.xlsm` into a standalone Python script that generates a complete washout matrix for product transitions.

## Generated Files

### Main Script
- **`washout_matrix_generator.py`** - The main Python script that replicates the VBA logic

### Output
- **`Washout_Matrix_Output.xlsx`** - Generated Excel file with two sheets:
  - **Washout Matrix**: Complete matrix of all product transitions (12,996 rows)
  - **Product Properties**: Reference table of all 114 products with their properties

### Supporting Files
- **`product_properties_full.csv`** - CSV export of the product properties table
- **`washout_formulas.txt`** - Documentation of the Excel formulas used in decision logic
- **`verify_output.py`** - Script to verify and analyze the generated output

## Input Data Source

The script reads from the original Excel file:
- **File**: `Alkox Washout Matrix Product Wheel.xlsm`
- **Sheet**: `2021 Washout Decision_old`
- **Product Table**: Columns K-Q, Rows 3-116 (114 products)

### Product Properties
Each product has the following properties:
- **Description** - Product name/code
- **Class** - Product class (e.g., Etho, TDA, Alcohol-BF3)
- **Starting Amine** - Base amine used
- **Family** - Product family
- **Oxide** - Oxide type (EO, PO, MO)
- **Moles Ox** - Number of moles of oxide
- **Category** - Category classification (e.g., TAE MM, CD 2M, Berol)

## Washout Decision Logic

The script implements the following decision rules:

### 1. Washout Needed? (YES/NO)
NO washout is needed if ANY of these conditions are true:
- Same oxide, same amine, and moles increasing (with exceptions for specific categories)
- Moving to TDA product from another alcohol with same oxide
- Both products are Berol category
- Both are Etho with same amine and both > 2 moles (with category exceptions)
- Next product's starting amine is the previous product (raw material relationship)

Otherwise, YES - washout is needed.

### 2. Washout Type (HEEL WASH / FULL WASH)
HEEL WASH is used for:
- Specific category transitions (18D MM, CD MM, SVD MM, TAE MM, etc.)
- Same amine and oxide, going from high to low molecular weight
- Same amine, different oxide, low to high molecular weight
- Certain Etho products with specific category combinations

FULL WASH is used for all other cases requiring washout.

### 3. Acid Wash Needed? (YES/NO)
YES if any of these conditions:
- Specific products: ETHOMEEN C/25A, C/17, C/15
- 2M products (except ETHOMEEN C/12M) when previous is not 2M
- Propo to Etho transition
- ETHODUOMEEN T/13 when previous is not 2M
- All ARMOSTAT products

### 4. Duration Calculation
- **HEEL WASH**: 1.5 hours
- **FULL WASH**: 2 hours
- **Acid Wash**: +2 hours (added to base time)
- **No Wash**: 0 hours

### 5. Special Notes
Additional notes are generated for:
- Berol products (BF3 wash requirements)
- V700 unit specific rules
- Ethylan TD-100 campaign rules
- Post-Ethomeen 18/12 requirements
- ARMOSTAT FWR (Fresh Water Rinse) requirements
- C/15, C/17, C/25A acid wash instructions

## Usage

### Running the Script

```powershell
# Make sure you're in the correct directory
cd "L:\Washout Matrix"

# Run the main generator script
python washout_matrix_generator.py
```

### Output Statistics
The generated matrix contains:
- **Total transitions**: 12,996 (114 products × 114 products)
- **Washouts needed**: 12,484 (96.1%)
- **No washout**: 512 (3.9%)
- **HEEL WASH**: 900 (7.2%)
- **FULL WASH**: 11,584 (92.8%)
- **Acid wash required**: 1,579 (12.2%)

### Average Durations
- **Overall average**: 2.13 hours
- **When washout needed**: 2.22 hours
- **HEEL WASH average**: 1.80 hours
- **FULL WASH average**: 2.25 hours
- **With acid wash**: 3.94 hours

## Output Excel File Structure

### Sheet 1: Washout Matrix
Columns:
- `From_Product` - Starting product name
- `From_Class` - Starting product class
- `From_Amine` - Starting product amine
- `From_Oxide` - Starting product oxide
- `From_Moles` - Starting product moles
- `From_Category` - Starting product category
- `To_Product` - Destination product name
- `To_Class` - Destination product class
- `To_Amine` - Destination product amine
- `To_Oxide` - Destination product oxide
- `To_Moles` - Destination product moles
- `To_Category` - Destination product category
- `Washout_Needed` - YES or NO
- `Washout_Type` - HEEL WASH or FULL WASH (blank if no washout)
- `Acid_Wash` - YES or NO
- `Notes` - Special instructions and notes
- `Duration_Hours` - Estimated changeover time in hours

### Sheet 2: Product Properties
Reference table with all 114 products and their properties.

## Verification

To verify the output and see statistics:

```powershell
python verify_output.py
```

This will display:
- Total number of transitions
- Summary statistics by washout type
- Sample transitions
- Average duration statistics

## Dependencies

Required Python packages (already installed in the virtual environment):
- `pandas` - Data manipulation
- `openpyxl` - Excel file handling

## Notes

- The script creates a virtual environment at `.venv/` if one doesn't exist
- All intermediate files are saved in the same directory as the script
- The output Excel file is formatted with headers and proper column widths
- The script preserves all the original VBA logic including special cases and exceptions

## Maintenance

If new products are added to the Excel file:
1. Update the product table in the Excel file (columns K-Q in the '2021 Washout Decision_old' sheet)
2. Run the Python script again to regenerate the matrix
3. The script will automatically include all products found in rows 3-116

## Contact

For questions about the washout logic or product properties, refer to:
- **2021 Rules** sheet in the original Excel file for general rules
- **Formula Code** sheet for the original formula documentation
- **washout_formulas.txt** for the extracted Excel formulas
