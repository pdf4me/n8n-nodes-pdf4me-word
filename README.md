# n8n-nodes-pdf4me-word

This is an n8n community node that enables you to process Word documents with PDF4ME's powerful Word processing capabilities. It includes a suite of Word-focused actions such as adding text/image watermarks, extracting metadata, optimizing, comparing, splitting/merging, securing with passwords, updating TOC, replacing text, and updating headers/footers.

n8n is a fair-code licensed workflow automation platform.

## Table of Contents

- [Installation](#installation)
- [Operations](#operations)
- [Credentials](#credentials)
- [Usage](#usage)
- [Resources](#resources)
- [Version History](#version-history)

## Installation

### Community Nodes (Recommended)

For users on n8n v0.187+, you can install this node directly from the n8n Community Nodes panel in the n8n editor:

1. Open your n8n editor
2. Go to **Settings > Community Nodes**
3. Search for "n8n-nodes-pdf4me-word"
4. Click **Install**
5. Reload the editor

### Manual Installation

You can also install this node manually in a specific n8n project:

1. Navigate to your n8n installation directory
2. Run the following command:
   ```bash
   npm install n8n-nodes-pdf4me-word
   ```
3. Restart your n8n server

For Docker-based deployments, add the package to your package.json and rebuild the image:

```json
{
  "name": "n8n-custom",
  "version": "1.0.0",
  "description": "",
  "main": "index.js",
  "scripts": {
    "start": "n8n"
  },
  "dependencies": {
    "n8n": "^1.0.0",
    "n8n-nodes-pdf4me-word": "^0.8.0"
  }
}
```

## Operations

Below are the available Word operations exposed by this node:

- **Add Text Watermark To Word**: Add customizable text watermarks with font, color, rotation, orientation, and transparency controls.
- **Add Image Watermark To Word**: Add image watermarks with configurable scale, size, alignment, and transparency.
- **Extract Word Metadata**: Extract document properties and statistics from Word files.
- **Optimize Word Document**: Reduce file size and improve performance with configurable optimization levels.
- **Compare Word Documents**: Compare two Word documents and produce a redlined output highlighting differences.
- **Split Word Document**: Split by pages, sections, headings, or custom ranges into multiple documents.
- **Merge Word Documents**: Merge multiple Word files into a single document with merge options.
- **Delete Pages From Word**: Remove specified pages or ranges with optional pagination updates.
- **Secure Word Document**: Apply password protection and protection types (ReadOnly, CommentsOnly, FormsOnly).
- **Update Table of Contents**: Update TOC with heading levels, page numbers, and tab leaders.
- **Replace Text**: Find and replace text with search options, formatting, and regex.
- **Update Headers and Footers**: Update content for different page types (first, odd, even).

## Credentials

To use this node, you need a PDF4ME API key:

1. Sign up for a free account at [PDF4ME](https://portal.pdf4me.com/register)
2. Navigate to [API Keys](https://portal.pdf4me.com/api-keys) in your account
3. Generate a new API key
4. Add the API key to your n8n credentials:
   - In n8n, go to **Credentials > New**
   - Select **PDF4me API**
   - Enter your API key
   - Save the credentials

## Usage

### Basic Example: Add Watermark to Word Document

This example shows how to add a text watermark to a Word file:

1. **Input Node**: Use a node that provides a Word file (e.g., HTTP Request, Google Drive, etc.)
2. **PDF4me Word Node**: Configure with:
   - Operation: Add Text Watermark To Word
   - Word File Input Method: From Previous Node (Binary Data)
   - Binary Data Property Name: data
   - Watermark Text: CONFIDENTIAL
   - Orientation: Diagonal
   - Font Family: Arial
   - Font Size: 72
   - Font Color: #808080
   - Semi Transparent: true
   - Rotation: 45
   - Culture Name: en-US
   - Output File Name: word_with_watermark.docx
3. **Output**: The modified Word file with watermark added

### Advanced Example: Process Multiple Files

Process multiple Word files with custom watermarks:

1. **Loop Node**: Iterate over multiple files
2. **PDF4me Word Node**: Add watermarks to each file
3. **Save/Send**: Save the processed files or send them via email

### Action Details

Below are concise configuration notes for each action. In n8n, set the operation in `Pdf4me Word` and fill the highlighted fields. Binary I/O defaults to property name `data` unless you customize it.

#### Add Text Watermark To Word
- **Inputs**: `Word File`, `Watermark Text`, `Orientation`, `Font Family`, `Font Size`, `Font Color`, `Semi Transparent`, `Rotation`, `Culture Name`
- **Output**: Word file with text watermark
- **Tips**: Use large font (60–120) and semi-transparent gray for readable diagonal watermarks

#### Add Image Watermark To Word
- **Inputs**: `Word File`, `Image` (binary/base64/URL), `Scale/Size`, `Alignment`, `Transparency`
- **Output**: Word file with image watermark
- **Tips**: Prefer PNG with transparency; control size via `Scale` or explicit `Width/Height`

#### Extract Word Metadata
- **Inputs**: `Word File`
- **Output**: JSON with document properties and statistics
- **Tips**: Use this to drive routing decisions in workflows

#### Optimize Word Document
- **Inputs**: `Word File`, `Optimization Level` (e.g., Balanced, Maximum)
- **Output**: Optimized Word file
- **Tips**: Balanced works well for most cases; verify layout-sensitive docs after optimization

#### Compare Word Documents
- **Inputs**: `Original Word File`, `Revised Word File`, `Comparison Options` (author, granularity)
- **Output**: Redlined Word file with tracked differences
- **Tips**: Keep both inputs in consistent culture/locale for best diff fidelity

#### Split Word Document
- **Inputs**: `Word File`, `Split Mode` (Pages, Sections, Headings, Custom Range), `Ranges` (when applicable)
- **Output**: Array of Word files, one per split part
- **Tips**: For large splits, consider downstream batch save or compression

#### Merge Word Documents
- **Inputs**: `Files` (multiple Word inputs), `Merge Options` (order, section breaks)
- **Output**: Single merged Word file
- **Tips**: Ensure consistent page size/margins across sources to avoid layout jumps

#### Delete Pages From Word
- **Inputs**: `Word File`, `Pages/Ranges` (e.g., 1,3-5)
- **Output**: Word file with pages removed
- **Tips**: Enable pagination updates if your doc contains TOC or cross-references

#### Secure Word Document
- **Inputs**: `Word File`, `Password`, `Protection Type` (ReadOnly, CommentsOnly, FormsOnly)
- **Output**: Password-protected or restricted Word file
- **Tips**: Store passwords in n8n credentials or environment variables, not hardcoded strings

#### Update Table of Contents
- **Inputs**: `Word File`, `Heading Levels`, `Show Page Numbers`, `Tab Leader`
- **Output**: Word file with refreshed TOC
- **Tips**: Run after structural edits, splits/merges, or page deletions

#### Replace Text
- **Inputs**: `Word File`, `Find Text`, `Replace With`, `Options` (match case, whole word, regex), `Formatting` (optional)
- **Output**: Word file with replacements applied
- **Tips**: Test regex on small samples; consider culture settings for locale-specific rules

#### Update Headers and Footers
- **Inputs**: `Word File`, `Page Types` (First, Odd, Even), `Content` per section
- **Output**: Word file with updated headers/footers
- **Tips**: If using different odd/even pages, supply both variants explicitly

### Input Options

#### Word File Input Method
- **From Previous Node (Binary Data)**: Use Word files from previous nodes (most common)
- **Base64 Encoded String**: Provide Word content as base64 encoded string
- **Download from URL**: Download Word file directly from a web URL

#### Watermark Configuration
- **Watermark Text**: Text to display as watermark (e.g., CONFIDENTIAL, DRAFT, INTERNAL USE ONLY)
- **Orientation**: Horizontal, Vertical, Diagonal, or Upside-Down
- **Font Family**: Arial, Times New Roman, Courier New, Verdana, Calibri, Helvetica, Georgia, or Tahoma
- **Font Size**: 6-500 points
- **Font Color**: Hex color code (e.g., #808080 for gray)
- **Semi Transparent**: Whether the watermark should be semi-transparent (true/false)
- **Rotation**: Rotation angle from -360 to 360 degrees
- **Culture Name**: Culture name for document (e.g., en-US, fr-FR, de-DE)

#### Output Options
- **Output File Name**: Name for the processed Word file
- **Source Document Name**: Name of the original Word file (for reference)
- **Output Binary Data Name**: Custom name for the binary data in n8n output

## Resources

- [PDF4ME API Documentation](https://dev.pdf4me.com/apiv2/documentation/)
- [PDF4ME Portal](https://portal.pdf4me.com/)
- [n8n Documentation](https://docs.n8n.io/)
- [n8n Community](https://community.n8n.io/)

## Compatibility

- Minimum n8n version: 0.187.0
- Supported Word formats: .docx
- API: PDF4ME API v2

## Support

For issues and feature requests:
- GitHub Issues: [n8n-nodes-pdf4me-word](https://github.com/pdf4me/n8n-nodes-pdf4me-word/issues)
- PDF4ME Support: support@pdf4me.com
- n8n Community: [community.n8n.io](https://community.n8n.io/)

## Version History

### 0.8.0 (Current)
- Initial release of `n8n-nodes-pdf4me-word`
- Word operations: Add Text/Image Watermark, Extract Metadata, Optimize, Compare, Split, Merge, Delete Pages, Secure, Update TOC, Replace Text, Update Headers/Footers
- Multiple input methods (binary data, base64, URL)
- Support for culture names where applicable

## License

[MIT License](LICENSE.md)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Keywords

n8n, n8n-community-node-package, word, docx, watermark, pdf4me, office, documents, automation

## Author

PDF4me - https://pdf4me.com

## Acknowledgments

Built with [n8n](https://n8n.io/) and powered by [PDF4ME API](https://pdf4me.com/)
