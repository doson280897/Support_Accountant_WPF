#!/usr/bin/env python3
import re
import pdfplumber
import shutil
import argparse
from pathlib import Path

def load_patterns():
    date_patterns = [
        r'Ngày\s*(?:\([^)]*\))?\s*(\d{1,2})\s*tháng\s*(?:\([^)]*\))?\s*(\d{1,2})\s*năm\s*(?:\([^)]*\))?\s*(\d{4})',
        r'Ngày\s*tháng\s*năm/?\s*Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})',
        r'Ngày\s*lập:\s*(\d{1,2})/(\d{1,2})/(\d{4})',
        r'Ngày\s*\(Dated\)\s*:\s*(\d{1,2})/(\d{1,2})/(\d{4})'
    ]
    
    number_patterns = [
        r'Số\s*(?:\([^)]*\))?\s*:?\s*(\d+)',
        r'(\d{4})\s+(\d+)\s*\n?\s*Số\s*(?:\([^)]*\))?\s*:',
        r'(\d{8})\s*.*?Ngày\s*\d{1,2}\s*tháng\s*\d{1,2}\s*năm\s*\d{4}\s*Số\s*:',
        r'(?:(?:Số\s*hóa\s*đơn|Invoice\s*No)[:/\s]*(\d+)|(?:VAT\s*INVOICE\)?\s*(\d{3,8})\s*.*?(?:Số\s*hóa\s*đơn|Invoice\s*No)))',
        r'(?:INVOICE\)?\s*(\d{4,8})\s*.*?Số\s*\([^)]*\)\s*:|Số\s*\([^)]*\)\s*:.*?(\d{4,8}))',
        r'Mã\s*số\s*thuế\s*\(?Tax\s*code\)?\s*:\s*\d+\s+Số\s*hóa\s*đơn\s*\(?Invoice\s*No\.\)?\s*:\s*(\d+)',
    ]
    return date_patterns, number_patterns

def extract_date_and_number(pdf_path, date_patterns, number_patterns):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        if not text.strip():
            return None, None

        date_result = None
        for pattern in date_patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                try:
                    day, month, year = match.groups()
                except ValueError:
                    continue
                year_short = year[-2:] if len(year) == 4 else year
                date_result = f"{year_short}{month.zfill(2)}{day.zfill(2)}"
                break

        number_result = None
        for i, pattern in enumerate(number_patterns):
            flags = re.DOTALL | re.IGNORECASE if i >= 2 else re.IGNORECASE if i == 3 else 0
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                if i == 1:
                    number_result = match.group(2)
                elif i in [3, 4]:
                    number_result = match.group(1) if match.group(1) else match.group(2)
                else:
                    number_result = match.group(1)
                break

        return date_result, number_result

def get_unique_filename(base_path, filename):
    destination = base_path / filename
    if not destination.exists():
        return filename
    name_part, extension = filename.rsplit('.', 1)
    counter = 1
    while True:
        new_filename = f"{name_part} ({counter}).{extension}"
        if not (base_path / new_filename).exists():
            return new_filename
        counter += 1

def process_files(pdf_files, success_dir, failed_dir):
    date_patterns, number_patterns = load_patterns()
    success_count = 0
    failed_count = 0

    success_dir = Path(success_dir)
    failed_dir = Path(failed_dir)
    success_dir.mkdir(exist_ok=True)
    failed_dir.mkdir(exist_ok=True)

    for pdf_file in pdf_files:
        pdf_file = Path(pdf_file)
        try:
            date_result, number_result = extract_date_and_number(str(pdf_file), date_patterns, number_patterns)

            if date_result and number_result:
                base_filename = f"{date_result}_{number_result}.pdf"
                unique_filename = get_unique_filename(success_dir, base_filename)
                shutil.copy2(pdf_file, success_dir / unique_filename)
                success_count += 1
                print(f"PROGRESS: {pdf_file.name} -> SUCCESS")
            else:
                shutil.copy2(pdf_file, failed_dir / pdf_file.name)
                failed_count += 1
                print(f"PROGRESS: {pdf_file.name} -> FAILED")
        except Exception as e:
            shutil.copy2(pdf_file, failed_dir / pdf_file.name)
            failed_count += 1
            print(f"PROGRESS: {pdf_file.name} -> ERROR: {str(e)}")

        # flush so C# receives immediately
        import sys; sys.stdout.flush()

    print(f"SUMMARY: SUCCESS={success_count}, FAILED={failed_count}")

def main():
    parser = argparse.ArgumentParser(description="Batch process PDFs and extract date/number")
    parser.add_argument("-i", "--inputs", nargs="+", required=True, help="List of input PDF files")
    parser.add_argument("-s", "--success", required=True, help="Output folder for successful extractions")
    parser.add_argument("-f", "--failed", required=True, help="Output folder for failed extractions")
    args = parser.parse_args()

    process_files(args.inputs, args.success, args.failed)

if __name__ == "__main__":
    main()
