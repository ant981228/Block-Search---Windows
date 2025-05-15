# Block Search

A sophisticated Word document search and content transfer utility designed specifically for debate research.

## Features

### Document Search
- Fast, as-you-type searching of Word documents
- Context-aware search that can include file paths in search results
- Prefix-based searching to limit results to specific folders
- View documents in their original context with hierarchy preservation
- Multiple sorting options (name, date modified, date created, size)

### Document Handling
- One-click document content transfer to:
  - Clipboard
  - Closed target document
  - Open Word document at cursor position or document end
- Preview documents before selecting
- View documents in their original document context
- Preserved formatting, including highlighting and background colors

### Document Splitting
- Split large Word documents by heading levels
- Preserve document hierarchy and relationships
- Output as individual files or ZIP archive
- Uses a template system for consistent formatting

### User Experience
- Global hotkey activation (default: Ctrl+Space)
- System tray integration
- Comprehensive keyboard navigation
- Status bar feedback
- Customizable UI

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+Space | Show/activate application (configurable) |
| Up/Down | Navigate through search results |
| Left Arrow | Show document in its original context |
| Right Arrow | Return from context view to search results |
| Enter | Select document (use default paste mode) |
| Ctrl+Enter | Select document (use alternate paste mode) |
| Shift+Enter | Preview document without selecting |
| Ctrl+T | Select (closed) target document |
| Ctrl+Shift+T | Clear (closed) target document |
| Ctrl+P | Toggle between paste modes (cursor/end) |
| Ctrl+Shift+P | Toggle including path names in search |
| F1 | Show help dialog |
| F5 | Refresh open documents list |
| Escape | Hide application window |
| Alt+F4, Ctrl+Q | Quit application |

## Installation

1. Ensure you have Python 3.6+ installed
2. Install required packages: `pip install -r requirements.txt`
3. Run the application: `python BlockSearch-Windows.py`

## Usage Tips

### Basic Search
Type search terms in the search box to find matching documents. Results update as you type.

### Using Prefixes
Configure prefixes to limit searches to specific folders:
1. Set up prefixes in Search Settings → Prefix Configuration → Manage Prefixes
2. Use format: `[prefix] [search terms]`
3. Example: `cb 2ac` searches for "2ac" only in folders assigned to "cb" prefix

### Document Context
Press Left Arrow on any search result to see it in the context of its original document.
Navigate with Up/Down arrows, select with Enter, and press Right Arrow to return to search.

### Splitting Documents
1. Open Document Tools → Split Document by Headings
2. Select input document and template (optional)
3. Choose heading level to split at and output options
4. Select output location
5. Click Process Document

## Development

This application is built with:
- Python 3.6+
- PyQt6 for the user interface
- python-docx for Word document handling
- win32com for Word automation

Contributions are welcome!