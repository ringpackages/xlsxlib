# XLSXLib Documentation

## Overview

XLSXLib is a library for creating Microsoft Excel (.xlsx) files using the Ring programming language. It generates fully compatible Office Open XML (ECMA-376) files that can be opened in Microsoft Excel, LibreOffice Calc, Google Sheets, and other spreadsheet applications.

## Features

- **Multiple Worksheets** - Create workbooks with many sheets
- **Cell Operations** - Set text, numbers, dates, and formulas
- **Styling** - Fonts, colors, backgrounds, borders, alignment
- **Number Formats** - Currency, percentage, date, time, custom formats
- **Column/Row Formatting** - Custom widths and heights
- **Merged Cells** - Combine cells horizontally and vertically
- **Auto Filter** - Add filter dropdowns to data ranges
- **Freeze Panes** - Keep headers visible while scrolling
- **Images** - Embed PNG, JPG, BMP, and GIF images
- **Formulas** - Full Excel formula support
- **No Dependencies** - Pure Ring implementation, no external libraries

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
  - [Simple Spreadsheet](#simple-spreadsheet)
  - [Quick Export Function](#quick-export-function)
  - [Using Styles](#using-styles)
- [API Reference](#api-reference)
  - [ExcelWriter Class](#excelwriter-class)
  - [Document Properties](#document-properties)
  - [Sheet Management](#sheet-management)
  - [Cell Operations](#cell-operations)
  - [Bulk Data Operations](#bulk-data-operations)
  - [Styling](#styling)
  - [Column and Row Formatting](#column-and-row-formatting)
  - [Merged Cells](#merged-cells)
  - [Auto Filter](#auto-filter)
  - [Freeze Panes](#freeze-panes)
  - [Images](#images)
  - [Saving](#saving)
- [Quick Functions](#quick-functions)
- [Constants](#constants)
- [Color Reference](#color-reference)

---

## Installation

	ringpm install xlsxlib from ringpackages

---

## Quick Start

### Simple Spreadsheet

```ring
load "xlsxlib.ring"

excel = new ExcelWriter()
excel.addSheet("Data")
excel.setCell(1, 1, "Hello")
excel.setCell(1, 2, "World")
excel.setCell(2, 1, 42)
excel.save("output.xlsx")
```

### Quick Export Function

```ring
load "xlsxlib.ring"

data = [
    ["Name", "Age", "City"],
    ["Ahmed", 25, "Cairo"],
    ["Sara", 30, "Riyadh"]
]

quickExcel("output.xlsx", data, "People")
```

### Using Styles

```ring
load "xlsxlib.ring"

excel = new ExcelWriter()
excel.addSheet("Report")

# Create a header style
headerStyle = excel.createStyle([
    :bold = true,
    :bgColor = "4472C4",
    :fontColor = "FFFFFF",
    :align = "center"
])

# Apply style to cells
excel.setCellWithStyle(1, 1, "Product", headerStyle)
excel.setCellWithStyle(1, 2, "Price", headerStyle)

excel.setCell(2, 1, "Laptop")
excel.setCell(2, 2, 999.99)

excel.save("report.xlsx")
```

---

## API Reference

### ExcelWriter Class

#### Constructor

```ring
excel = new ExcelWriter()
```

Creates a new Excel workbook.

---

### Document Properties

| Method | Description |
|--------|-------------|
| `setAuthor(author)` | Set document author |
| `setTitle(title)` | Set document title |
| `setCompany(company)` | Set company name |

**Example:**
```ring
excel.setAuthor("John Doe")
excel.setTitle("Sales Report 2025")
excel.setCompany("ABC Corporation")
```

---

### Sheet Management

| Method | Description |
|--------|-------------|
| `addSheet(name)` | Add a new worksheet |
| `selectSheet(index)` | Select sheet by index (1-based) |
| `selectSheetByName(name)` | Select sheet by name |
| `getSheetCount()` | Get number of sheets |

**Example:**
```ring
excel.addSheet("Sales")
excel.addSheet("Expenses")
excel.selectSheet(1)        # Select first sheet
excel.selectSheetByName("Expenses")  # Select by name
? excel.getSheetCount()     # Output: 2
```

---

### Cell Operations

| Method | Description |
|--------|-------------|
| `setCell(row, col, value)` | Set cell value (auto-detects type) |
| `setCellWithStyle(row, col, value, styleId)` | Set cell with specific style |
| `setFormula(row, col, formula)` | Set cell formula (without =) |
| `setNumber(row, col, value)` | Set numeric value |
| `setDate(row, col, year, month, day)` | Set date value |

**Example:**
```ring
excel.setCell(1, 1, "Text")          # String
excel.setCell(1, 2, 42)              # Number
excel.setCell(1, 3, 3.14159)         # Decimal
excel.setFormula(1, 4, "A1&B1")      # Formula (without =)
excel.setFormula(2, 1, "SUM(A1:C1)") # Sum formula
excel.setDate(3, 1, 2025, 12, 25)    # Date
```

---

### Bulk Data Operations

| Method | Description |
|--------|-------------|
| `exportList(data, startRow, startCol, hasHeader)` | Export 2D list to cells |
| `setRow(row, values, startCol)` | Set entire row |
| `setColumn(col, values, startRow)` | Set entire column |

**Example:**
```ring
# Export a 2D list with header formatting
data = [
    ["Name", "Score"],
    ["Alice", 95],
    ["Bob", 87]
]
excel.exportList(data, 1, 1, true)  # true = first row is header

# Set a row of values
excel.setRow(5, ["Total", 182], 1)

# Set a column of values
excel.setColumn(3, ["A", "B", "C"], 1)
```

---

### Styling

#### Creating Styles

```ring
styleId = excel.createStyle(options)
```

**Available Options:**

| Option | Type | Description |
|--------|------|-------------|
| `:bold` | Boolean | Bold text |
| `:italic` | Boolean | Italic text |
| `:fontName` | String | Font name (e.g., "Arial") |
| `:fontSize` | Number | Font size in points |
| `:fontColor` | String | Font color (name or hex) |
| `:bgColor` | String | Background color (name or hex) |
| `:align` | String | Horizontal alignment: "left", "center", "right" |
| `:valign` | String | Vertical alignment: "top", "center", "bottom" |
| `:wrap` | Boolean | Wrap text in cell |
| `:border` | String | Border style (see below) |
| `:borderColor` | String | Border color (name or hex) |
| `:numberFormat` | Number | Number format ID (see constants) |

**Example:**
```ring
# Create a custom style
myStyle = excel.createStyle([
    :bold = true,
    :italic = true,
    :fontName = "Arial",
    :fontSize = 14,
    :fontColor = "FF0000",
    :bgColor = "FFFF00",
    :align = "center",
    :valign = "center",
    :wrap = true,
    :border = "thin",
    :borderColor = "000000"
])

excel.setCellWithStyle(1, 1, "Styled Cell", myStyle)
```

#### Pre-defined Style Methods

| Method | Description |
|--------|-------------|
| `createHeaderStyle()` | Blue background, white bold text, centered |
| `createCurrencyStyle()` | Currency number format |
| `createPercentageStyle()` | Percentage number format |
| `createDateStyle()` | Date number format |

**Example:**
```ring
headerStyle = excel.createHeaderStyle()
currencyStyle = excel.createCurrencyStyle()

excel.setCellWithStyle(1, 1, "Price", headerStyle)
excel.setCellWithStyle(2, 1, 99.99, currencyStyle)
```

---

### Column and Row Formatting

| Method | Description |
|--------|-------------|
| `setColumnWidth(col, width)` | Set column width (in characters) |
| `setRowHeight(row, height)` | Set row height (in points) |
| `autoFitColumn(col)` | Auto-fit column width (approximate) |

**Example:**
```ring
excel.setColumnWidth(1, 20)   # Column A = 20 characters wide
excel.setColumnWidth(2, 15)   # Column B = 15 characters wide
excel.setRowHeight(1, 30)     # Row 1 = 30 points tall
```

---

### Merged Cells

```ring
excel.mergeCells(row1, col1, row2, col2)
```

Merges cells from (row1, col1) to (row2, col2).

**Example:**
```ring
excel.setCell(1, 1, "Title")
excel.mergeCells(1, 1, 1, 4)  # Merge A1:D1

excel.setCell(3, 1, "Subtitle")
excel.mergeCells(3, 1, 5, 1)  # Merge A3:A5 (vertical)
```

---

### Auto Filter

```ring
excel.setAutoFilter(row1, col1, row2, col2)
```

Adds filter dropdowns to the specified range.

**Example:**
```ring
# Add headers and data rows...
# Enable filter on the data range
excel.setAutoFilter(1, 1, 10, 3)  # A1:C10
```

---

### Freeze Panes

| Method | Description |
|--------|-------------|
| `freezePanes(row, col)` | Freeze above row and left of column |
| `freezeTopRow()` | Freeze first row only |
| `freezeFirstColumn()` | Freeze first column only |

**Example:**
```ring
excel.freezeTopRow()        # Freeze row 1
excel.freezeFirstColumn()   # Freeze column A
excel.freezePanes(2, 2)     # Freeze row 1 AND column A
```

---

### Images

```ring
excel.addImage(imagePath, row, col, width, height)
```

**Parameters:**
- `imagePath` - Path to image file (PNG, JPG, BMP, GIF)
- `row`, `col` - Cell position for top-left corner
- `width` - Width in cells (default: 5)
- `height` - Height in cells (default: 5)

**Example:**
```ring
excel.addImage("logo.png", 1, 1, 4, 3)      # A1, 4 cells wide, 3 cells tall
excel.addImage("chart.jpg", 10, 1, 8, 6)    # A10, 8x6 cells
```

**Supported Formats:** PNG, JPG, JPEG, BMP, GIF

---

### Saving

```ring
result = excel.save(filename)
```

Returns `true` if successful, `false` otherwise.

---

## Quick Functions

| Function | Description |
|----------|-------------|
| `quickExcel(filename, data, sheetName)` | Create simple Excel from 2D list |
| `listToExcel(data, filename, sheetName, hasHeader)` | Export list with optional header |
| `listsToExcel(dataList, filename)` | Create multiple sheets |

---

## Constants

### Alignment
- `EXCEL_ALIGN_LEFT`, `EXCEL_ALIGN_CENTER`, `EXCEL_ALIGN_RIGHT`
- `EXCEL_ALIGN_TOP`, `EXCEL_ALIGN_MIDDLE`, `EXCEL_ALIGN_BOTTOM`

### Border Styles
- `EXCEL_BORDER_THIN`, `EXCEL_BORDER_MEDIUM`, `EXCEL_BORDER_THICK`
- `EXCEL_BORDER_DOUBLE`, `EXCEL_BORDER_DASHED`, `EXCEL_BORDER_DOTTED`

### Number Formats
- `EXCEL_FORMAT_GENERAL` (0)
- `EXCEL_FORMAT_NUMBER` (1)
- `EXCEL_FORMAT_CURRENCY` (164)
- `EXCEL_FORMAT_PERCENTAGE` (10)
- `EXCEL_FORMAT_DATE` (14)
- `EXCEL_FORMAT_TIME` (21)
- `EXCEL_FORMAT_DATETIME` (22)

---

## Color Reference

### Named Colors
black, white, red, green, blue, yellow, orange, purple, gray/grey, navy, teal, maroon, silver, lime, aqua, fuchsia, olive

### Hex Colors
```ring
excel.createStyle([:bgColor = "FF5733"])      # Without #
excel.createStyle([:fontColor = "#3498DB"])   # With #
```

---

