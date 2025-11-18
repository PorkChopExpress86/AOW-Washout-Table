# Alkox Washout Matrix - Python Tool

## Quick Start

### Generate and Verify Washout Matrix (Recommended)
```powershell
cd "L:\Washout Matrix"
python main.py
```

This will automatically:
1. Generate the complete washout matrix
2. Verify the output and show statistics

Output will be created at: `output\Washout_Matrix_Output.xlsx`

### Or Run Scripts Individually
```powershell
python scripts\washout_matrix_generator.py
python scripts\verify_output.py
```

---

## Project Structure

```
L:\Washout Matrix\
├── main.py                                   # Main entry point (recommended)
├── Alkox Washout Matrix Product Wheel.xlsm  # Original Excel file (input)
├── scripts\                                  # Python scripts
│   ├── washout_matrix_generator.py          # Main generator script
│   ├── verify_output.py                     # Validate output
│   ├── compare_outputs.py                   # Compare with original Excel
│   └── add_simple_sheet.py                  # Add simplified sheet
├── output\                                   # Generated files
│   ├── Washout_Matrix_Output.xlsx           # Main output file
│   ├── product_properties.csv               # Product data
│   └── product_properties_full.csv          # Complete product data
├── documentation\                            # Documentation
│   ├── README.md                            # Full user guide
│   ├── PROJECT_SUMMARY.md                   # Project overview
│   └── washout_formulas.txt                 # Excel formulas reference
└── .venv\                                   # Python virtual environment
```

---

## Main Output File

**`output\Washout_Matrix_Output.xlsx`** contains 3 sheets:

1. **Simple Matrix** - From Product, To Product, Duration (Hours)
2. **Washout Matrix** - Complete details (washout type, acid wash, notes)
3. **Product Properties** - Reference table of all 114 products

---

## Usage

### 1. Generate and Verify (One Command)
```powershell
python main.py
```

### 2. Compare with Original Excel (Optional)
```powershell
python scripts\compare_outputs.py
```

---

## Output Summary

- **12,996 transitions** (114 products × 114 products)
- **96.1%** require washout
- **Average duration**: 2.13 hours
- **HEEL WASH**: 1.5 hours base
- **FULL WASH**: 2.0 hours base
- **Acid Wash**: +2.0 hours

---

## For More Information

See `documentation\README.md` for:
- Detailed logic explanation
- Product properties structure
- Washout decision rules
- Maintenance guidelines

See `documentation\PROJECT_SUMMARY.md` for:
- Complete project overview
- Technical details
- Validation results
- Future enhancements
