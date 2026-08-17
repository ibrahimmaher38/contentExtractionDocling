from pathlib import Path
from docling.document_converter import DocumentConverter

# ---------------------------------------------------------
# 1. Define the input and output locations
# ---------------------------------------------------------

input_files = [
    Path("input/Account Closure-BRD-HDB V0.2.docx"),
    Path("input/Account Closure-BRD-HDB V0.2.pdf"),
    Path("input/DOCX_TestPage.docx"),
    Path("input/DOCX_TestPage.pdf"),
]

output_dir = Path("output")
output_dir.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------
# 2. Create Docling converter
# ---------------------------------------------------------

converter = DocumentConverter()

# ---------------------------------------------------------
# 3. Convert and export each document
# ---------------------------------------------------------

for input_file in input_files:
    result = converter.convert(input_file)

    # The conversion result contains the DoclingDocument,
    # which is Docling's structured representation of the document.
    doc = result.document

    # Get the input filename without its extension
    file_name = input_file.stem

    # Add the original file type to distinguish DOCX and PDF
    file_type = input_file.suffix[1:].upper()

    output_name = f"{file_name}_{file_type}"

    # -----------------------------------------------------
    # Export to Markdown

    markdown_content = doc.export_to_markdown()

    (output_dir / f"{output_name}.md").write_text(
        markdown_content,
        encoding="utf-8"
    )

    # -----------------------------------------------------
    # Export to JSON

    doc.save_as_json(
        output_dir / f"{output_name}.json"
    )

    # -----------------------------------------------------
    # Export to HTML

    doc.save_as_html(
        output_dir / f"{output_name}.html"
    )

# ---------------------------------------------------------
# 4. Print completion message
# ---------------------------------------------------------

print("Document processed successfully!")
print(f"Output files saved to: {output_dir}")