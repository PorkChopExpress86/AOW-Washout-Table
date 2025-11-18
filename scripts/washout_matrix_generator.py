"""
Alkox Washout Matrix Generator
================================
This script replicates the VBA macro logic from the Excel file to generate
a washout matrix for product transitions.

The script:
1. Reads product properties from the product table (columns K-Q, rows 3-116)
2. Generates all from/to product combinations
3. Applies washout decision logic to determine:
   - Whether a washout is needed (YES/NO)
   - Washout type (HEEL WASH / FULL WASH)
   - Whether acid wash is needed (YES/NO)
   - Special notes and instructions
   - Estimated changeover duration
4. Outputs results to an Excel file
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime


class WashoutMatrixGenerator:
    def __init__(self, input_file):
        """Initialize with the path to the input Excel file"""
        self.input_file = input_file
        self.products_df = None

    def load_products(self):
        """Load product properties from the Excel file"""
        print("Loading product properties...")

        # Load the workbook
        wb = openpyxl.load_workbook(self.input_file, data_only=True)
        ws = wb["2021 Washout Decision_old"]

        # Read product table from K3:Q116
        products = []
        for row in range(3, 117):
            product = {
                "Description": ws.cell(row=row, column=11).value,  # K
                "Class": ws.cell(row=row, column=12).value,  # L
                "Starting_Amine": ws.cell(row=row, column=13).value,  # M
                "Family": ws.cell(row=row, column=14).value,  # N
                "Oxide": ws.cell(row=row, column=15).value,  # O
                "Moles_Ox": ws.cell(row=row, column=16).value,  # P
                "Category": ws.cell(row=row, column=17).value,  # Q
            }

            # Only add if Description exists
            if product["Description"]:
                # Convert Moles_Ox to numeric if possible
                try:
                    if isinstance(product["Moles_Ox"], str):
                        # Handle cases like "12 PO, 5 EO" - take first number
                        parts = product["Moles_Ox"].split(",")
                        product["Moles_Ox"] = float(parts[0].split()[0])
                    else:
                        product["Moles_Ox"] = float(product["Moles_Ox"])
                except:
                    product["Moles_Ox"] = 0

                products.append(product)

        wb.close()

        self.products_df = pd.DataFrame(products)
        print(f"Loaded {len(self.products_df)} products")
        return self.products_df

    def decide_washout_needed(self, prod1, prod2):
        """
        Determine if washout is needed based on the formula:
        =IF(OR(
            AND(F3=F9,E3=E9,G3<=G9,NOT(OR(H3="18D MM",H3="CD MM",H3="SVD MM",AND(H3="TAE MM",OR(G3=5,G3=16))))),
            AND(D9="TDA",F3=F9,G3<=G9,OR(D3="Nonyl Phenol",D3="Alcohol C9-C11",D3="Alcohol C10-12")),
            AND(H3="Berol",H9="Berol"),
            AND(D3="Etho",E3=E9,G3>2,G9>2,NOT(OR(H3="18D MM",H3="CD MM",H3="SVD MM",AND(H3="TAE MM",OR(G3=5,G3=16))))),
            E9=C3
        ),"NO","YES")

        Where:
        C3 = prod1 Description, D3 = prod1 Class, E3 = prod1 Starting_Amine,
        F3 = prod1 Oxide, G3 = prod1 Moles_Ox, H3 = prod1 Category
        C9 = prod2 Description, D9 = prod2 Class, E9 = prod2 Starting_Amine,
        F9 = prod2 Oxide, G9 = prod2 Moles_Ox, H9 = prod2 Category
        """

        # Condition 1: Same oxide, same amine, lower/equal moles, except certain categories
        cond1 = (
            prod1["Oxide"] == prod2["Oxide"]
            and prod1["Starting_Amine"] == prod2["Starting_Amine"]
            and prod1["Moles_Ox"] <= prod2["Moles_Ox"]
            and not (
                prod1["Category"] in ["18D MM", "CD MM", "SVD MM"]
                or (prod1["Category"] == "TAE MM" and prod1["Moles_Ox"] in [5, 16])
            )
        )

        # Condition 2: Moving to TDA with same oxide
        cond2 = (
            prod2["Class"] == "TDA"
            and prod1["Oxide"] == prod2["Oxide"]
            and prod1["Moles_Ox"] <= prod2["Moles_Ox"]
            and prod1["Class"] in ["Nonyl Phenol", "Alcohol C9-C11", "Alcohol C10-12"]
        )

        # Condition 3: Both Berol
        cond3 = prod1["Category"] == "Berol" and prod2["Category"] == "Berol"

        # Condition 4: Etho with same amine, both > 2 moles
        cond4 = (
            prod1["Class"] == "Etho"
            and prod1["Starting_Amine"] == prod2["Starting_Amine"]
            and prod1["Moles_Ox"] > 2
            and prod2["Moles_Ox"] > 2
            and not (
                prod1["Category"] in ["18D MM", "CD MM", "SVD MM"]
                or (prod1["Category"] == "TAE MM" and prod1["Moles_Ox"] in [5, 16])
            )
        )

        # Condition 5: Next product's amine is the previous product (raw material relationship)
        cond5 = prod2["Starting_Amine"] == prod1["Description"]

        # If any condition is true, NO washout needed
        if cond1 or cond2 or cond3 or cond4 or cond5:
            return "NO"
        else:
            return "YES"

    def decide_washout_type(self, prod1, prod2, washout_needed):
        """
        Determine washout type (HEEL WASH vs FULL WASH)
        Returns empty string if no washout needed
        """
        if washout_needed == "NO":
            return ""

        cat1 = prod1["Category"]
        cat2 = prod2["Category"]

        # All conditions that require HEEL WASH
        heel_wash_conditions = [
            # 18D MM transitions
            cat1 == "18D MM" and cat2 in ["18D MM", "18D 2M", "SVD MM", "TAE MM"],
            # 18D 2M transitions
            cat1 == "18D 2M" and cat2 in ["E-9", "TAE 2M", "TAE MM"],
            # CD 2M transitions
            cat1 == "CD 2M"
            and cat2
            in ["18D 2M", "18D MM", "E-9", "SVD 2M", "SVD MM", "TAE 2M", "TAE MM"],
            # CD MM transitions
            cat1 == "CD MM" and cat2 in ["CD 2M", "CD MM", "E-9", "SVD MM", "TAE MM"],
            # SVD 2M transitions
            cat1 == "SVD 2M"
            and cat2
            in ["CD 2M", "CD MM", "E-9", "TAE 2M", "TAE MM", "18D MM", "SVD 2M"],
            # SVD MM transitions
            cat1 == "SVD MM"
            and cat2
            in ["CD 2M", "CD MM", "E-9", "TAE 2M", "TAE MM", "18D MM", "SVD 2M"],
            # TAE 2M transitions
            cat1 == "TAE 2M" and cat2 in ["18D MM", "E-9", "SVD 2M", "SVD MM"],
            # TAE MM transitions
            cat1 == "TAE MM" and cat2 in ["18D MM", "E-9", "SVD MM", "TAE 2M"],
            # Same amine and oxide, going from high to low MW (not Nonyl Phenol)
            (
                prod1["Starting_Amine"] == prod2["Starting_Amine"]
                and prod1["Oxide"] == prod2["Oxide"]
                and prod1["Moles_Ox"] >= prod2["Moles_Ox"]
                and prod1["Starting_Amine"] != "Nonyl Phenol"
            ),
            # Same amine, different oxide, low to high MW
            (
                prod1["Starting_Amine"] == prod2["Starting_Amine"]
                and prod1["Oxide"] != prod2["Oxide"]
                and prod1["Moles_Ox"] <= prod2["Moles_Ox"]
            ),
            # Etho with same amine, both > 2 moles, special categories
            (
                prod1["Class"] == "Etho"
                and prod1["Starting_Amine"] == prod2["Starting_Amine"]
                and prod1["Moles_Ox"] > 2
                and prod2["Moles_Ox"] > 2
                and (
                    cat1 in ["18D MM", "CD MM", "SVD MM"]
                    or (
                        cat1 == "TAE MM"
                        and (prod1["Moles_Ox"] == 5 or prod2["Moles_Ox"] == 16)
                    )
                )
            ),
        ]

        if any(heel_wash_conditions):
            return "HEEL WASH"
        else:
            return "FULL WASH"

    def decide_acid_wash(self, prod1, prod2):
        """
        Determine if acid wash is needed
        Formula: =IF(OR(
            C9="ETHOMEEN C/25A",C9="ETHOMEEN C/17",C9="ETHOMEEN C/15",
            AND(OR(D9="Etho",D9="Propo"),G9=2,G3<>2,NOT(C9="ETHOMEEN C/12M")),
            AND(D9="Propo",D3="Etho"),
            AND(C9="ETHODUOMEEN T/13",G3<>2),
            LEFT(C9,8)="ARMOSTAT"
        ),"YES","NO")
        """

        # Specific products that need acid wash
        if prod2["Description"] in ["ETHOMEEN C/25A", "ETHOMEEN C/17", "ETHOMEEN C/15"]:
            return "YES"

        # 2M products (except ETHOMEEN C/12M) when previous is not 2M
        if (
            prod2["Class"] in ["Etho", "Propo"]
            and prod2["Moles_Ox"] == 2
            and prod1["Moles_Ox"] != 2
            and prod2["Description"] != "ETHOMEEN C/12M"
        ):
            return "YES"

        # Propo to Etho transition
        if prod2["Class"] == "Propo" and prod1["Class"] == "Etho":
            return "YES"

        # ETHODUOMEEN T/13 when previous is not 2M
        if prod2["Description"] == "ETHODUOMEEN T/13" and prod1["Moles_Ox"] != 2:
            return "YES"

        # ARMOSTAT products
        if prod2["Description"].startswith("ARMOSTAT"):
            return "YES"

        return "NO"

    def get_notes(self, prod1, prod2, washout_needed, washout_type, acid_wash):
        """Generate notes based on various conditions"""
        notes = []

        # Berol products
        if prod2["Class"] == "Alcohol-BF3":
            notes.append(
                "Before any Berol campaing, do 1 Full Wash and 1 Soft Water Wash with 50lb of BF3. No wash in between consecutive Berol batches"
            )

        # V700 specific
        if prod2["Description"] in ["ETHOMEEN T/15", "ETHOMEEN T/25"]:
            notes.append(
                "If making the product in V700, no need for heel wash in between batches"
            )

        # Ethylan TD-100
        if prod2["Description"] == "Ethylan TD-100":
            notes.append(
                "During campaings of Ethylan TD-100, wash every 3 batches in a row"
            )

        # After Ethomeen 18/12
        if prod1["Description"] == "Ethomeen 18/12":
            notes.append("After Ethomeen 18/12 do 2 full washes")

        # ARMOSTAT FWR note
        if prod2["Description"].startswith("ARMOSTAT"):
            notes.append("Only first batch needs a FWR")

        # C/15, C/17, C/25A acid wash note
        if prod2["Description"] in ["ETHOMEEN C/15", "ETHOMEEN C/17", "ETHOMEEN C/25A"]:
            notes.append("Perform an acid wash before batch (add acetic to full wash)")

        if not notes:
            return "------"

        return " | ".join(notes)

    def calculate_duration(self, washout_needed, washout_type, acid_wash):
        """
        Calculate changeover duration in hours
        Formula: =IF(E24="NO",0,IF(F24="HEEL WASH",1.5,2))+IF(G24="NO",0,2)
        """
        if washout_needed == "NO":
            return 0

        # Base wash time
        if washout_type == "HEEL WASH":
            base_time = 1.5
        else:  # FULL WASH
            base_time = 2

        # Add acid wash time if needed
        if acid_wash == "YES":
            acid_time = 2
        else:
            acid_time = 0

        return base_time + acid_time

    def generate_matrix(self):
        """Generate the complete washout matrix for all product combinations"""
        print("Generating washout matrix...")

        if self.products_df is None:
            self.load_products()

        results = []
        total_combinations = len(self.products_df) ** 2

        # Generate all from/to combinations
        for idx1, prod1 in self.products_df.iterrows():
            for idx2, prod2 in self.products_df.iterrows():
                # Decide washout
                washout_needed = self.decide_washout_needed(prod1, prod2)
                washout_type = self.decide_washout_type(prod1, prod2, washout_needed)
                acid_wash = self.decide_acid_wash(prod1, prod2)
                notes = self.get_notes(
                    prod1, prod2, washout_needed, washout_type, acid_wash
                )
                duration = self.calculate_duration(
                    washout_needed, washout_type, acid_wash
                )

                result = {
                    "From_Product": prod1["Description"],
                    "From_Class": prod1["Class"],
                    "From_Amine": prod1["Starting_Amine"],
                    "From_Oxide": prod1["Oxide"],
                    "From_Moles": prod1["Moles_Ox"],
                    "From_Category": prod1["Category"],
                    "To_Product": prod2["Description"],
                    "To_Class": prod2["Class"],
                    "To_Amine": prod2["Starting_Amine"],
                    "To_Oxide": prod2["Oxide"],
                    "To_Moles": prod2["Moles_Ox"],
                    "To_Category": prod2["Category"],
                    "Washout_Needed": washout_needed,
                    "Washout_Type": washout_type,
                    "Acid_Wash": acid_wash,
                    "Notes": notes,
                    "Duration_Hours": duration,
                }

                results.append(result)

        print(f"Generated {len(results)} washout transitions")
        return pd.DataFrame(results)

    def export_to_excel(self, output_file, include_product_properties=True):
        """Export the washout matrix to an Excel file"""
        print(f"Exporting to {output_file}...")

        # Generate the matrix
        matrix_df = self.generate_matrix()

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            # Write the main washout matrix
            matrix_df.to_excel(writer, sheet_name="Washout Matrix", index=False)

            # Write product properties if requested
            if include_product_properties:
                self.products_df.to_excel(
                    writer, sheet_name="Product Properties", index=False
                )

        # Format the Excel file
        self._format_excel(output_file)

        print(f"Successfully exported to {output_file}")

    def _format_excel(self, filename):
        """Apply formatting to the Excel file"""
        wb = openpyxl.load_workbook(filename)

        # Format Washout Matrix sheet
        ws = wb["Washout Matrix"]

        # Header formatting
        header_fill = PatternFill(
            start_color="4472C4", end_color="4472C4", fill_type="solid"
        )
        header_font = Font(bold=True, color="FFFFFF")

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Freeze header row
        ws.freeze_panes = "A2"

        wb.save(filename)


def main():
    """Main execution function"""
    print("=" * 60)
    print("Alkox Washout Matrix Generator")
    print("=" * 60)
    print()

    # File paths
    input_file = r"L:\Washout Matrix\Alkox Washout Matrix Product Wheel.xlsm"
    output_file = r"L:\Washout Matrix\Washout_Matrix_Output.xlsx"

    # Create generator
    generator = WashoutMatrixGenerator(input_file)

    # Load products
    products = generator.load_products()
    print(f"\nProduct properties sample:")
    print(products.head())
    print()

    # Generate and export matrix
    generator.export_to_excel(output_file, include_product_properties=True)

    print()
    print("=" * 60)
    print("Process completed successfully!")
    print(f"Output file: {output_file}")
    print("=" * 60)


if __name__ == "__main__":
    main()
