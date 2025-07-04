#!/usr/bin/env python3
"""
Markdown to DOCX Converter Demo
===============================

This script demonstrates the different conversion approaches available
in the markdown_to_word project.
"""

import os
import sys
import time
from pathlib import Path

# Import all converters
from markdown_converter import MarkdownConverter
from markdown_to_docx_converter import MarkdownToDocxConverter
from hybrid_converter import HybridMarkdownConverter, AdvancedMarkdownConverter

def ensure_output_dir():
    """Ensure output directory exists"""
    output_dir = Path("outputs")
    output_dir.mkdir(exist_ok=True)
    return output_dir

def convert_with_timing(converter, name, input_file, output_file):
    """Convert and measure time taken"""
    print(f"\n--- {name} ---")
    print(f"Converting: {input_file} -> {output_file}")
    
    start_time = time.time()
    try:
        if hasattr(converter, 'convert_markdown_to_docx'):
            converter.convert_markdown_to_docx(input_file, output_file)
        else:
            converter.convert(input_file, output_file)
        
        end_time = time.time()
        print(f"✅ Success! Time taken: {end_time - start_time:.2f} seconds")
        return True
    except Exception as e:
        end_time = time.time()
        print(f"❌ Error: {e}")
        print(f"Time taken: {end_time - start_time:.2f} seconds")
        return False

def main():
    """Main demonstration function"""
    print("=== Markdown to DOCX Converter Demo ===\n")
    
    # Ensure output directory exists
    output_dir = ensure_output_dir()
    
    # Test files
    test_files = [
        "test_basic.md",
        "test_complex.md", 
        "test_comprehensive.md"
    ]
    
    # Check if test files exist
    missing_files = [f for f in test_files if not Path(f).exists()]
    if missing_files:
        print(f"❌ Missing test files: {missing_files}")
        print("Please ensure test files exist in the current directory.")
        return
    
    # Initialize converters
    converters = [
        (MarkdownConverter(), "Basic Converter (md2docx_python)"),
        (MarkdownToDocxConverter(), "Custom Converter (python-docx)"),
        (HybridMarkdownConverter(), "Hybrid Converter (pypandoc + python-docx)"),
        (AdvancedMarkdownConverter(), "Advanced Converter (full custom parsing)")
    ]
    
    results = {}
    
    # Test each converter with each file
    for test_file in test_files:
        print(f"\n{'='*60}")
        print(f"Testing with: {test_file}")
        print('='*60)
        
        file_stem = Path(test_file).stem
        results[test_file] = {}
        
        for converter, name in converters:
            # Generate output filename
            converter_suffix = name.split()[0].lower()
            output_file = output_dir / f"{file_stem}_{converter_suffix}.docx"
            
            # Convert and record result
            success = convert_with_timing(converter, name, test_file, str(output_file))
            results[test_file][name] = success
    
    # Print summary
    print(f"\n{'='*60}")
    print("CONVERSION SUMMARY")
    print('='*60)
    
    for test_file, file_results in results.items():
        print(f"\n{test_file}:")
        for converter_name, success in file_results.items():
            status = "✅ Success" if success else "❌ Failed"
            print(f"  {converter_name}: {status}")
    
    # Print recommendations
    print(f"\n{'='*60}")
    print("RECOMMENDATIONS")
    print('='*60)
    
    print("""
1. **Hybrid Converter** - Best for most use cases
   - Uses pypandoc for comprehensive markdown support
   - Enhanced with python-docx for better formatting
   - Handles complex elements like footnotes, tables, etc.

2. **Advanced Converter** - Best for custom control
   - Full custom parsing with python-docx
   - Complete control over formatting
   - Good for specialized requirements

3. **Custom Converter** - Good for simple documents
   - Pure python-docx implementation
   - Good for basic markdown elements
   - Lightweight and fast

4. **Basic Converter** - Simplest approach
   - Uses md2docx_python library
   - Minimal configuration required
   - Good for quick conversions
    """)
    
    print(f"\nOutput files saved to: {output_dir.absolute()}")

def usage_examples():
    """Show usage examples"""
    print("\n=== USAGE EXAMPLES ===")
    
    examples = [
        ("Basic conversion", """
from markdown_converter import MarkdownConverter
converter = MarkdownConverter()
converter.convert_markdown_to_docx("input.md", "output.docx")
        """),
        
        ("Custom conversion with template", """
from markdown_to_docx_converter import MarkdownToDocxConverter
converter = MarkdownToDocxConverter(template_path="template.docx")
converter.convert("input.md", "output.docx")
        """),
        
        ("Hybrid conversion (recommended)", """
from hybrid_converter import HybridMarkdownConverter
converter = HybridMarkdownConverter()
converter.convert("input.md", "output.docx")
        """),
        
        ("Advanced conversion with full control", """
from hybrid_converter import AdvancedMarkdownConverter
converter = AdvancedMarkdownConverter()
converter.convert("input.md", "output.docx")
        """)
    ]
    
    for title, code in examples:
        print(f"\n{title}:")
        print(code)

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--examples":
        usage_examples()
    else:
        main()