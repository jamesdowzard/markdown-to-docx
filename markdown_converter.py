from md2docx_python.src.md2docx_python import markdown_to_word

class MarkdownConverter:
    def __init__(self):
        pass # No explicit initialization needed for markdown_to_word

    def convert_markdown_to_docx(self, markdown_file_path, output_docx_path):
        """
        Converts a Markdown file to a DOCX file.
        """
        try:
            markdown_to_word(markdown_file_path, output_docx_path)
            print(f"Successfully converted '{markdown_file_path}' to '{output_docx_path}'")
        except Exception as e:
            print(f"Error converting '{markdown_file_path}': {e}")

if __name__ == "__main__":
    converter = MarkdownConverter()
    
    # Convert the basic test file
    converter.convert_markdown_to_docx(
        "test_basic.md",
        "outputs/test_basic.docx"
    )

    # Convert the complex test file
    converter.convert_markdown_to_docx(
        "test_complex.md",
        "outputs/test_complex.docx"
    )