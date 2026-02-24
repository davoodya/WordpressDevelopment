#!/usr/bin/env python3
"""
Simple test script - creates a small test CSV and converts it.
"""

from pathlib import Path
import csv


def create_test_csv():
    """Create a simple test CSV with Persian text and duplicates."""
    
    test_file = Path(r"h:\Repo\WordpressDevelopment\Products-Price-Exporter\test_sample.csv")
    
    test_data = [
        ["نام محصول", "قیمت", "دسته‌بندی"],
        ["آیفون 14", "50000000", "موبایل"],
        ["سامسونگ گلکسی", "30000000", "موبایل"],
        ["  آیفون 14  ", "51000000", "موبایل"],  # Duplicate with spaces
        ["آیپد پرو", "40000000", "تبلت"],
        ["IPHONE 14", "52000000", "موبایل"],  # Duplicate different case
        ["لپتاپ HP Laptop", "25000000", "کامپیوتر"],
        ["", "", ""],  # Empty row
        ["ماوس Gaming Mouse", "500000", "لوازم جانبی"],
        ["لپتاپ hp laptop", "26000000", "کامپیوتر"],  # Duplicate mixed case
    ]
    
    with open(test_file, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(test_data)
    
    print(f"Created test CSV: {test_file}")
    return test_file


if __name__ == "__main__":
    csv_file = create_test_csv()
    print(f"\nYou can now run the converter manually with:")
    print(f'python csv-to-excel-inpu.py')
    print(f"\nOr use this path when prompted:")
    print(f"{csv_file}")
