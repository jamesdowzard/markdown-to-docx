# Comprehensive Markdown Test Document

This document tests all possible markdown elements for conversion to Microsoft Word format.

## Basic Text Formatting

**Bold text** using double asterisks
*Italic text* using single asterisks
***Bold and italic*** using triple asterisks
~~Strikethrough text~~ using tildes
`inline code` using backticks

Super^script^ and sub~script~ text (if supported)

## Headings Test

# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6

## Lists

### Unordered Lists
- Item 1
- Item 2
  - Nested item 2.1
  - Nested item 2.2
    - Deeply nested item 2.2.1
    - Deeply nested item 2.2.2
- Item 3

### Ordered Lists
1. First item
2. Second item
   1. Nested ordered item 2.1
   2. Nested ordered item 2.2
      1. Deeply nested item 2.2.1
      2. Deeply nested item 2.2.2
3. Third item

### Task Lists
- [x] Completed task
- [ ] Incomplete task
- [x] Another completed task
  - [ ] Nested incomplete task
  - [x] Nested completed task

### Mixed Lists
1. Ordered item 1
   - Unordered nested item
   - Another unordered nested item
2. Ordered item 2
   1. Nested ordered item
   2. Another nested ordered item

## Blockquotes

> This is a simple blockquote.
> It can span multiple lines.

> This is a blockquote with **bold** and *italic* text.

> Nested blockquotes:
> > This is nested inside another blockquote.
> > > This is triple nested.

## Code Blocks

### Inline Code
Here is some `inline code` in a sentence.

### Code Blocks with Language
```python
def hello_world():
    print("Hello, world!")
    return "success"

# This is a comment
for i in range(10):
    print(f"Number: {i}")
```

```javascript
function greetUser(name) {
    console.log(`Hello, ${name}!`);
    return true;
}

// Call the function
greetUser("World");
```

```bash
#!/bin/bash
echo "This is a bash script"
ls -la
cd /home/user
```

### Plain Code Block
```
This is a plain code block
without syntax highlighting
Multiple lines supported
```

## Links

### Inline Links
Visit [Google](https://www.google.com) for search.
Check out [GitHub](https://github.com) for code repositories.

### Reference Links
This is a [reference link][1] to Google.
This is another [reference link][google] using a label.

[1]: https://www.google.com
[google]: https://www.google.com "Google Homepage"

### Automatic Links
<https://www.example.com>
<user@example.com>

## Images

### Inline Images
![Alt text](https://via.placeholder.com/150x100 "Image Title")

### Reference Images
![Reference image][image1]

[image1]: https://via.placeholder.com/200x150 "Reference Image"

## Tables

### Simple Table
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |

### Table with Alignment
| Left Align | Center Align | Right Align |
|:-----------|:------------:|------------:|
| Left       | Center       | Right       |
| Text       | Text         | Text        |
| More       | More         | More        |

### Complex Table
| Feature | Description | Status | Priority |
|---------|-------------|--------|----------|
| **Headers** | Support for H1-H6 | ✅ Complete | High |
| *Lists* | Ordered and unordered | ✅ Complete | High |
| `Code` | Inline and blocks | ✅ Complete | Medium |
| Links | All types | 🟡 Partial | Medium |
| Images | With alt text | 🟡 Partial | Low |
| Tables | With alignment | ✅ Complete | Medium |

## Horizontal Rules

Here's a horizontal rule:

---

Another style:

***

And another:

___

## Footnotes

Here's a sentence with a footnote[^1].

This sentence has another footnote[^note].

[^1]: This is the first footnote.
[^note]: This is a named footnote with more content.
    It can span multiple lines and include **formatting**.

## Definition Lists

Term 1
: Definition for term 1

Term 2
: Definition for term 2
: Additional definition for term 2

Complex Term
: This is a complex definition that includes **bold** text,
  *italic* text, and even `code`.

## HTML Elements (if supported)

<div>This is raw HTML</div>

<span style="color: red;">Red text using HTML</span>

<table>
<tr><td>HTML Table</td><td>Cell 2</td></tr>
</table>

## Special Characters and Escaping

\*This text has escaped asterisks\*
\# This is not a heading
\[This is not a link\](example.com)

## Line Breaks

This line has two spaces at the end  
So this line should be on a new line.

This paragraph has a manual line break.

## Emphasis Combinations

**Bold with *italic* inside**
*Italic with **bold** inside*
***All bold and italic***
~~Strikethrough with **bold** inside~~
`Code with **bold** (should not work)`

## Extended Markdown Features

### Abbreviations (if supported)
*[HTML]: HyperText Markup Language
*[CSS]: Cascading Style Sheets

The HTML and CSS specifications are maintained by W3C.

### Highlight (if supported)
==This text should be highlighted==

### Keyboard Keys (if supported)
Press <kbd>Ctrl</kbd> + <kbd>C</kbd> to copy.

### Mathematical Expressions (if supported)
Inline math: $E = mc^2$

Block math:
$$
\int_{-\infty}^{\infty} e^{-x^2} dx = \sqrt{\pi}
$$

## Metadata (Front Matter)

```yaml
---
title: "Comprehensive Markdown Test"
author: "Test Author"
date: "2024-01-01"
description: "A comprehensive test of all markdown elements"
keywords: ["markdown", "docx", "conversion", "test"]
---
```

## Summary

This document contains examples of:
- [x] All basic markdown elements
- [x] Extended markdown features  
- [x] Complex nested structures
- [x] Various formatting combinations
- [x] Edge cases and special characters
- [ ] Some features may not be supported in all converters

**Total elements tested:** 50+ markdown features

*Document created for comprehensive markdown to DOCX conversion testing.*