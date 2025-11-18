"""
Main entry point for Alkox Washout Matrix Generator
Generates the washout matrix and verifies the output
"""

import subprocess
import sys
from pathlib import Path


def run_script(script_name, description):
    """Run a Python script and handle errors"""
    print("\n" + "=" * 70)
    print(f"{description}")
    print("=" * 70)

    script_path = Path(__file__).parent / "scripts" / script_name

    try:
        result = subprocess.run(
            [sys.executable, str(script_path)],
            check=True,
            capture_output=False,
            text=True,
        )
        print(f"\n✓ {description} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n✗ Error running {script_name}")
        print(f"Exit code: {e.returncode}")
        return False
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        return False


def main():
    """Main execution function"""
    print("=" * 70)
    print("ALKOX WASHOUT MATRIX GENERATOR")
    print("=" * 70)
    print("\nThis script will:")
    print("1. Generate the complete washout matrix")
    print("2. Verify the output and show statistics")
    print()

    # Step 1: Generate the matrix
    success = run_script(
        "washout_matrix_generator.py", "Step 1: Generating Washout Matrix"
    )

    if not success:
        print("\n✗ Failed to generate washout matrix. Exiting.")
        sys.exit(1)

    # Step 2: Verify the output
    success = run_script("verify_output.py", "Step 2: Verifying Output")

    if not success:
        print("\n⚠ Warning: Verification had issues, but output was generated.")

    # Final summary
    print("\n" + "=" * 70)
    print("PROCESS COMPLETE")
    print("=" * 70)
    print("\nOutput file location:")
    output_path = Path(__file__).parent / "output" / "Washout_Matrix_Output.xlsx"
    print(f"  {output_path}")
    print("\nThe Excel file contains 3 sheets:")
    print("  1. Simple Matrix - From/To products with duration")
    print("  2. Washout Matrix - Complete details")
    print("  3. Product Properties - Reference table")
    print("\n" + "=" * 70)


if __name__ == "__main__":
    main()
