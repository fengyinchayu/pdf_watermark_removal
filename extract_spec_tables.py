from is_micro_test import is_micro_test
import pdfplumber

def extract_spec_tables(pdf_file):
    general_rows = []
    micro_rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue

                    if len(row) >= 3:
                        row_data = {
                            "Characteristic": row[0].strip(),
                            "Specification": row[1].strip(),
                            "Method": row[2].strip()
                        }

                        if is_micro_test(row_data["Characteristic"]):
                            micro_rows.append(row_data)
                        else:
                            general_rows.append(row_data)

    return general_rows, micro_rows
