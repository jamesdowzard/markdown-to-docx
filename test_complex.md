# Complex Markdown Document

## Introduction
This document demonstrates the conversion of a complex Markdown file to a Word document, including various formatting elements.

## Headings

### Level 3 Heading
#### Level 4 Heading
##### Level 5 Heading
###### Level 6 Heading

## Text Formatting

**Bold text**, *italic text*, ***bold and italic text***, ~~strikethrough~~, and `inline code`.

## Lists

### Unordered List
* Item 1
  * Nested Item 1.1
    * Nested Item 1.1.1
* Item 2

### Ordered List
1. First item
2. Second item
   1. Nested ordered item 2.1
   2. Nested ordered item 2.2
3. Third item

## Blockquotes

> This is a blockquote.
> It can span multiple lines.
> > Nested blockquote.

## Code Blocks

```javascript
function greet(name) {
  console.log(`Hello, ${name}!`);
}

greet("World");
```

## Links

Visit [Google](https://www.google.com).

## Images

![Alt text for image](https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png "Google Logo")

## Tables

### Simple Table

| Header A | Header B | Header C |
| :------- | :------: | -------: |
| Left     | Center   | Right    |
| Cell 1   | Cell 2   | Cell 3   |

### Table with more content

| Feature        | Description                                     | Status   |
| :------------- | :---------------------------------------------- | :------- |
| **Headings**   | Supports H1-H6                                  | Implemented |
| *Lists*        | Ordered and unordered, nested                   | Implemented |
| `Code Blocks`  | Syntax highlighting (if supported by Word)      | Implemented |
| Links          | Inline and reference-style                      | Implemented |
| Images         | Inline images with alt text                     | Implemented |
| Tables         | Basic and complex with alignment                | Implemented |
| Blockquotes    | Single and nested                               | Implemented |

## Horizontal Rule

---

## Footnotes (if supported by md2docx-python)

Here is some text with a footnote.[^1]

[^1]: This is the footnote content.

## Task List (if supported by md2docx-python)

- [x] Task 1 (completed)
- [ ] Task 2 (pending)

## Definition List (if supported by md2docx-python)

Term 1
: Definition 1

Term 2
: Definition 2.1
: Definition 2.2

## Conclusion

This document serves as a comprehensive test for Markdown to Word conversion capabilities.
