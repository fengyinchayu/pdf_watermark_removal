from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream, NameObject
from pypdf.generic import TextStringObject
from pypdf.generic import DecodedStreamObject

def remove_watermark(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        content = page.get_contents()
        content_obj = ContentStream(content, reader)

        new_operations = []

        for operands, operator in content_obj.operations:
            # Remove suspected watermark operators
            if operator == b"Do":
                # Skip XObject drawing (often watermark)
                continue
            if operator == b"gs":
                continue
            new_operations.append((operands, operator))

        content_obj.operations = new_operations

        new_stream = DecodedStreamObject()
        new_stream.set_data(content_obj.get_data())

        page[NameObject("/Contents")] = new_stream
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    print("Watermark filtering attempt completed.")

# ---------------- MAIN ----------------
if __name__ == "__main__":
    INPUT_PDF = "COA-Organic Astragalus Root Extract 10 to1 ORHQ-20250430.pdf"
    OUTPUT_PDF = "SN-COA-Organic Astragalus Root Extract 10 to1.pdf"

    remove_watermark(INPUT_PDF, OUTPUT_PDF)