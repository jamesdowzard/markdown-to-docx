# Markdown to DOCX Converter

A comprehensive Python system for converting Markdown files to Microsoft Word (DOCX) format with full support for all markdown elements.

## 🚀 Quick Start

```bash
# Install dependencies
pip install python-docx pypandoc
brew install pandoc  # macOS

# Run the demo
python converter_demo.py
```

## 📋 Features

### ✅ Fully Supported Markdown Elements
- **Headings** (H1-H6)
- **Text formatting** (bold, italic, code, strikethrough)
- **Lists** (ordered, unordered, nested, task lists)
- **Code blocks** with syntax highlighting
- **Tables** with alignment and borders
- **Blockquotes** (single and nested)
- **Links** (inline and reference-style)
- **Images** with alt text
- **Footnotes** and references
- **Horizontal rules**
- **Definition lists**

## 🔧 Available Converters

### 1. Hybrid Converter (⭐ Recommended)
Best for complex documents with all markdown features.
```python
from hybrid_converter import HybridMarkdownConverter
converter = HybridMarkdownConverter()
converter.convert("input.md", "output.docx")
```

### 2. Advanced Converter
Best for custom formatting requirements.
```python
from hybrid_converter import AdvancedMarkdownConverter
converter = AdvancedMarkdownConverter()
converter.convert("input.md", "output.docx")
```

### 3. Basic Converter
Best for quick, simple conversions.
```python
from markdown_converter import MarkdownConverter
converter = MarkdownConverter()
converter.convert_markdown_to_docx("input.md", "output.docx")
```

## 📊 Performance Comparison

| Converter | Speed | Feature Support | Customization |
|-----------|-------|-----------------|---------------|
| Hybrid | Moderate | Excellent | High |
| Advanced | Fast | Excellent | Very High |
| Basic | Fast | Good | Limited |

## 📁 Project Structure

```
markdown_to_word/
├── README.md                    # This file
├── converter_demo.py            # Demo script
├── hybrid_converter.py          # Hybrid & advanced converters
├── markdown_converter.py        # Basic converter
├── markdown_to_docx_converter.py # Custom converter
├── test_basic.md               # Basic test file
├── test_complex.md             # Complex test file
├── test_comprehensive.md       # Comprehensive test file
└── outputs/                    # Generated DOCX files
```

## 🧪 Testing

Run the comprehensive test suite:
```bash
python converter_demo.py
```

This will test all converters with all test files and provide a detailed comparison.

## 📖 Documentation

For detailed documentation, see the comprehensive comments and examples in the source code files which include:
- Complete API reference
- Style template creation
- Advanced configuration options
- Troubleshooting guide
- Development guidelines

## 🛠️ Requirements

- Python 3.6+
- python-docx
- pypandoc
- pandoc (system dependency)

## 📄 License

This project is open source and available under the MIT License.