/*
    XLSXLib Demo - Excel File Creation in Ring
*/

load "xlsxlib.ring"

? "=============================================="
? "   XLSXLib Demo - Excel Creation in Ring"
? "=============================================="
? ""

# Demo 1: Simple Spreadsheet
? "Demo 1: Creating a simple spreadsheet..."
excel = new ExcelWriter()
excel.setTitle("Simple Spreadsheet")
excel.setAuthor("Ring Programmer")

excel.addSheet("Data")
excel.setCell(1, 1, "Name")
excel.setCell(1, 2, "Age")
excel.setCell(1, 3, "City")
excel.setCell(2, 1, "Ahmed")
excel.setCell(2, 2, 25)
excel.setCell(2, 3, "Cairo")
excel.setCell(3, 1, "Sara")
excel.setCell(3, 2, 30)
excel.setCell(3, 3, "Riyadh")

if excel.save("demo1_simple.xlsx")
    ? "  Created: demo1_simple.xlsx"
else
    ? "  FAILED: demo1_simple.xlsx"
ok

# Demo 2: Multiple Sheets
? "Demo 2: Creating workbook with multiple sheets..."
excel = new ExcelWriter()
excel.setTitle("Multi-Sheet Workbook")

excel.addSheet("Sales")
excel.setCell(1, 1, "Product")
excel.setCell(1, 2, "Revenue")
excel.setCell(2, 1, "Laptops")
excel.setCell(2, 2, 50000)
excel.setCell(3, 1, "Phones")
excel.setCell(3, 2, 35000)

excel.addSheet("Expenses")
excel.setCell(1, 1, "Category")
excel.setCell(1, 2, "Amount")
excel.setCell(2, 1, "Salaries")
excel.setCell(2, 2, 25000)
excel.setCell(3, 1, "Rent")
excel.setCell(3, 2, 5000)

excel.addSheet("Summary")
excel.setCell(1, 1, "Total Revenue")
excel.setCell(1, 2, 85000)
excel.setCell(2, 1, "Total Expenses")
excel.setCell(2, 2, 30000)
excel.setCell(3, 1, "Profit")
excel.setCell(3, 2, 55000)

if excel.save("demo2_multiple_sheets.xlsx")
    ? "  Created: demo2_multiple_sheets.xlsx"
else
    ? "  FAILED: demo2_multiple_sheets.xlsx"
ok

# Demo 3: Formulas
? "Demo 3: Creating spreadsheet with formulas..."
excel = new ExcelWriter()
excel.setTitle("Formulas Demo")

excel.addSheet("Calculations")
excel.setCell(1, 1, "Item")
excel.setCell(1, 2, "Quantity")
excel.setCell(1, 3, "Price")
excel.setCell(1, 4, "Total")

excel.setCell(2, 1, "Apples")
excel.setCell(2, 2, 10)
excel.setCell(2, 3, 2.50)
excel.setFormula(2, 4, "B2*C2")

excel.setCell(3, 1, "Oranges")
excel.setCell(3, 2, 15)
excel.setCell(3, 3, 1.75)
excel.setFormula(3, 4, "B3*C3")

excel.setCell(4, 1, "Bananas")
excel.setCell(4, 2, 20)
excel.setCell(4, 3, 0.99)
excel.setFormula(4, 4, "B4*C4")

excel.setCell(6, 1, "Grand Total:")
excel.setFormula(6, 4, "SUM(D2:D4)")

if excel.save("demo3_formulas.xlsx")
    ? "  Created: demo3_formulas.xlsx"
else
    ? "  FAILED: demo3_formulas.xlsx"
ok

# Demo 4: Styling
? "Demo 4: Creating spreadsheet with styles..."
excel = new ExcelWriter()
excel.setTitle("Styled Spreadsheet")

excel.addSheet("Styled Data")

# Create styles
headerStyle = excel.createStyle([
    :bold = true, 
    :bgColor = "4472C4", 
    :fontColor = "FFFFFF",
    :align = "center"
])

currencyStyle = excel.createStyle([
    :numberFormat = 164,
    :align = "right"
])

highlightStyle = excel.createStyle([
    :bold = true,
    :bgColor = "FFFF00"
])

# Apply styles
excel.setCellWithStyle(1, 1, "Product", headerStyle)
excel.setCellWithStyle(1, 2, "Price", headerStyle)
excel.setCellWithStyle(1, 3, "Status", headerStyle)

excel.setCell(2, 1, "Laptop")
excel.setCellWithStyle(2, 2, 999.99, currencyStyle)
excel.setCell(2, 3, "In Stock")

excel.setCell(3, 1, "Phone")
excel.setCellWithStyle(3, 2, 599.99, currencyStyle)
excel.setCellWithStyle(3, 3, "Low Stock", highlightStyle)

excel.setCell(4, 1, "Tablet")
excel.setCellWithStyle(4, 2, 399.99, currencyStyle)
excel.setCell(4, 3, "In Stock")

if excel.save("demo4_styled.xlsx")
    ? "  Created: demo4_styled.xlsx"
else
    ? "  FAILED: demo4_styled.xlsx"
ok

# Demo 5: Column Widths and Row Heights
? "Demo 5: Creating spreadsheet with custom dimensions..."
excel = new ExcelWriter()
excel.setTitle("Custom Dimensions")

excel.addSheet("Dimensions")
excel.setCell(1, 1, "Wide Column")
excel.setCell(1, 2, "Normal")
excel.setCell(1, 3, "Another Wide Column")

excel.setCell(2, 1, "This column is extra wide")
excel.setCell(2, 2, "Normal width")
excel.setCell(2, 3, "Also wide")

excel.setColumnWidth(1, 30)
excel.setColumnWidth(3, 25)
excel.setRowHeight(1, 30)

if excel.save("demo5_dimensions.xlsx")
    ? "  Created: demo5_dimensions.xlsx"
else
    ? "  FAILED: demo5_dimensions.xlsx"
ok

# Demo 6: Merged Cells
? "Demo 6: Creating spreadsheet with merged cells..."
excel = new ExcelWriter()
excel.setTitle("Merged Cells")

excel.addSheet("Report")

# Merge cells for title
excel.setCell(1, 1, "Quarterly Sales Report")
excel.mergeCells(1, 1, 1, 4)

excel.setCell(3, 1, "Region")
excel.setCell(3, 2, "Q1")
excel.setCell(3, 3, "Q2")
excel.setCell(3, 4, "Total")

excel.setCell(4, 1, "North")
excel.setCell(4, 2, 10000)
excel.setCell(4, 3, 12000)
excel.setFormula(4, 4, "B4+C4")

excel.setCell(5, 1, "South")
excel.setCell(5, 2, 8000)
excel.setCell(5, 3, 9500)
excel.setFormula(5, 4, "B5+C5")

if excel.save("demo6_merged.xlsx")
    ? "  Created: demo6_merged.xlsx"
else
    ? "  FAILED: demo6_merged.xlsx"
ok

# Demo 7: Auto Filter
? "Demo 7: Creating spreadsheet with auto filter..."
excel = new ExcelWriter()
excel.setTitle("Auto Filter Demo")

excel.addSheet("Employees")
excel.setCell(1, 1, "Name")
excel.setCell(1, 2, "Department")
excel.setCell(1, 3, "Salary")

excel.setCell(2, 1, "Ahmed")
excel.setCell(2, 2, "Engineering")
excel.setCell(2, 3, 75000)

excel.setCell(3, 1, "Sara")
excel.setCell(3, 2, "Marketing")
excel.setCell(3, 3, 65000)

excel.setCell(4, 1, "Omar")
excel.setCell(4, 2, "Engineering")
excel.setCell(4, 3, 70000)

excel.setCell(5, 1, "Fatima")
excel.setCell(5, 2, "HR")
excel.setCell(5, 3, 60000)

excel.setAutoFilter(1, 1, 5, 3)

if excel.save("demo7_autofilter.xlsx")
    ? "  Created: demo7_autofilter.xlsx"
else
    ? "  FAILED: demo7_autofilter.xlsx"
ok

# Demo 8: Freeze Panes
? "Demo 8: Creating spreadsheet with frozen panes..."
excel = new ExcelWriter()
excel.setTitle("Freeze Panes")

excel.addSheet("Data")

# Header row
excel.setCell(1, 1, "ID")
excel.setCell(1, 2, "Name")
excel.setCell(1, 3, "Value")

# Data rows
for i = 2 to 50
    excel.setCell(i, 1, i - 1)
    excel.setCell(i, 2, "Item " + (i - 1))
    excel.setCell(i, 3, (i - 1) * 100)
next

# Freeze the header row
excel.freezeTopRow()

if excel.save("demo8_freeze_panes.xlsx")
    ? "  Created: demo8_freeze_panes.xlsx"
else
    ? "  FAILED: demo8_freeze_panes.xlsx"
ok

# Demo 9: Quick Export
? "Demo 9: Using quick export function..."
data = [
    ["Product", "Category", "Price"],
    ["Laptop", "Electronics", 999],
    ["Mouse", "Accessories", 25],
    ["Keyboard", "Accessories", 75],
    ["Monitor", "Electronics", 350]
]

quickExcel("demo9_quick.xlsx", data, "Products")
? "  Created: demo9_quick.xlsx"

# Demo 10: Multiple sheets from list
? "Demo 10: Creating multiple sheets from list..."
dataList = [
    ["Q1 Sales", [
        ["Product", "Revenue"],
        ["Laptops", 50000],
        ["Phones", 35000]
    ], true],
    ["Q2 Sales", [
        ["Product", "Revenue"],
        ["Laptops", 55000],
        ["Phones", 40000]
    ], true]
]

listsToExcel(dataList, "demo10_multi_sheet.xlsx")
? "  Created: demo10_multi_sheet.xlsx"

# Demo 11: Images
? "Demo 11: Creating spreadsheet with images..."
excel = new ExcelWriter()
excel.setTitle("Images Demo")

excel.addSheet("Images")
excel.setCell(1, 1, "Product Catalog with Images")
excel.mergeCells(1, 1, 1, 5)

excel.setCell(3, 1, "Product 1:")

# Try to add images if they exist
if fexists("images/test1.png")
    excel.addImage("images/test1.png", 4, 1, 4, 4)
    excel.setCell(9, 1, "Image: images/test1.png")
else
    excel.setCell(4, 1, "(images/test1.png not found)")
ok

excel.setCell(3, 6, "Product 2:")
if fexists("images/test2.jpg")
    excel.addImage("images/test2.jpg", 4, 6, 4, 4)
    excel.setCell(9, 6, "Image: images/test2.jpg")
else
    excel.setCell(4, 6, "(images/test2.jpg not found)")
ok

excel.setCell(11, 1, "Product 3:")
if fexists("images/test3.bmp")
    excel.addImage("images/test3.bmp", 12, 1, 4, 4)
    excel.setCell(17, 1, "Image: images/test3.bmp")
else
    excel.setCell(12, 1, "(images/test3.bmp not found)")
ok

if excel.save("demo11_images.xlsx")
    ? "  Created: demo11_images.xlsx"
else
    ? "  FAILED: demo11_images.xlsx"
ok

# Demo 12: Comprehensive Report
? "Demo 12: Creating comprehensive sales report..."
excel = new ExcelWriter()
excel.setTitle("Sales Report 2024")
excel.setAuthor("Sales Department")
excel.setCompany("ABC Corporation")

# Summary Sheet
excel.addSheet("Summary")

titleStyle = excel.createStyle([:bold = true, :fontSize = 16, :align = "center"])
headerStyle = excel.createHeaderStyle()
currencyStyle = excel.createStyle([:align = "right"])
totalStyle = excel.createStyle([:bold = true, :bgColor = "E2EFDA"])

excel.setCellWithStyle(1, 1, "Annual Sales Report 2024", titleStyle)
excel.mergeCells(1, 1, 1, 4)

excel.setCellWithStyle(3, 1, "Quarter", headerStyle)
excel.setCellWithStyle(3, 2, "Revenue", headerStyle)
excel.setCellWithStyle(3, 3, "Expenses", headerStyle)
excel.setCellWithStyle(3, 4, "Profit", headerStyle)

excel.setCell(4, 1, "Q1")
excel.setCellWithStyle(4, 2, 250000, currencyStyle)
excel.setCellWithStyle(4, 3, 180000, currencyStyle)
excel.setFormula(4, 4, "B4-C4")

excel.setCell(5, 1, "Q2")
excel.setCellWithStyle(5, 2, 280000, currencyStyle)
excel.setCellWithStyle(5, 3, 190000, currencyStyle)
excel.setFormula(5, 4, "B5-C5")

excel.setCell(6, 1, "Q3")
excel.setCellWithStyle(6, 2, 310000, currencyStyle)
excel.setCellWithStyle(6, 3, 200000, currencyStyle)
excel.setFormula(6, 4, "B6-C6")

excel.setCell(7, 1, "Q4")
excel.setCellWithStyle(7, 2, 350000, currencyStyle)
excel.setCellWithStyle(7, 3, 220000, currencyStyle)
excel.setFormula(7, 4, "B7-C7")

excel.setCellWithStyle(9, 1, "Total", totalStyle)
excel.setFormula(9, 2, "SUM(B4:B7)")
excel.setFormula(9, 3, "SUM(C4:C7)")
excel.setFormula(9, 4, "SUM(D4:D7)")

excel.setColumnWidth(1, 15)
excel.setColumnWidth(2, 15)
excel.setColumnWidth(3, 15)
excel.setColumnWidth(4, 15)

excel.setAutoFilter(3, 1, 7, 4)
excel.freezeTopRow()

# Details Sheet
excel.addSheet("Details")

excel.setCellWithStyle(1, 1, "Product", headerStyle)
excel.setCellWithStyle(1, 2, "Category", headerStyle)
excel.setCellWithStyle(1, 3, "Units Sold", headerStyle)
excel.setCellWithStyle(1, 4, "Unit Price", headerStyle)
excel.setCellWithStyle(1, 5, "Total Revenue", headerStyle)

products = [
    ["Laptop Pro", "Electronics", 450, 1299],
    ["Laptop Basic", "Electronics", 800, 699],
    ["Wireless Mouse", "Accessories", 2500, 29],
    ["Mechanical Keyboard", "Accessories", 1200, 89],
    ["4K Monitor", "Electronics", 600, 449],
    ["USB Hub", "Accessories", 3000, 25],
    ["Webcam HD", "Electronics", 1500, 79],
    ["Headphones", "Accessories", 2000, 149]
]

row = 2
productsLen = len(products)
for i = 1 to productsLen
    p = products[i]
    excel.setCell(row, 1, p[1])
    excel.setCell(row, 2, p[2])
    excel.setCell(row, 3, p[3])
    excel.setCellWithStyle(row, 4, p[4], currencyStyle)
    excel.setFormula(row, 5, "C" + row + "*D" + row)
    row++
next

excel.setColumnWidth(1, 20)
excel.setColumnWidth(2, 15)
excel.setColumnWidth(3, 12)
excel.setColumnWidth(4, 12)
excel.setColumnWidth(5, 15)

excel.setAutoFilter(1, 1, row - 1, 5)
excel.freezeTopRow()

if excel.save("demo12_comprehensive.xlsx")
    ? "  Created: demo12_comprehensive.xlsx"
else
    ? "  FAILED: demo12_comprehensive.xlsx"
ok

# Demo 13: Border Styles
? "Demo 13: Creating spreadsheet with border styles..."
excel = new ExcelWriter()
excel.setTitle("Border Styles")

excel.addSheet("Borders")

thinBorder = excel.createStyle([:border = "thin"])
mediumBorder = excel.createStyle([:border = "medium"])
thickBorder = excel.createStyle([:border = "thick"])
doubleBorder = excel.createStyle([:border = "double"])

excel.setCell(1, 1, "Border Styles Demo")
excel.mergeCells(1, 1, 1, 3)

excel.setCellWithStyle(3, 1, "Thin Border", thinBorder)
excel.setCellWithStyle(3, 2, "Cell 1", thinBorder)
excel.setCellWithStyle(3, 3, "Cell 2", thinBorder)

excel.setCellWithStyle(5, 1, "Medium Border", mediumBorder)
excel.setCellWithStyle(5, 2, "Cell 1", mediumBorder)
excel.setCellWithStyle(5, 3, "Cell 2", mediumBorder)

excel.setCellWithStyle(7, 1, "Thick Border", thickBorder)
excel.setCellWithStyle(7, 2, "Cell 1", thickBorder)
excel.setCellWithStyle(7, 3, "Cell 2", thickBorder)

excel.setCellWithStyle(9, 1, "Double Border", doubleBorder)
excel.setCellWithStyle(9, 2, "Cell 1", doubleBorder)
excel.setCellWithStyle(9, 3, "Cell 2", doubleBorder)

excel.setColumnWidth(1, 18)
excel.setColumnWidth(2, 12)
excel.setColumnWidth(3, 12)

if excel.save("demo13_borders.xlsx")
    ? "  Created: demo13_borders.xlsx"
else
    ? "  FAILED: demo13_borders.xlsx"
ok

# Demo 14: Color Palette
? "Demo 14: Creating spreadsheet with colors..."
excel = new ExcelWriter()
excel.setTitle("Color Demo")

excel.addSheet("Colors")

colors = [
    "red", "green", "blue", "yellow", "orange", "purple",
    "navy", "teal", "maroon", "gray", "silver", "lime"
]

excel.setCell(1, 1, "Background Colors")
excel.setCell(1, 3, "Font Colors")

row = 3
colorsLen = len(colors)
for i = 1 to colorsLen
    color = colors[i]
    
    bgStyle = excel.createStyle([:bgColor = color])
    fontStyle = excel.createStyle([:fontColor = color, :bold = true])
    
    excel.setCellWithStyle(row, 1, color, bgStyle)
    excel.setCellWithStyle(row, 3, color, fontStyle)
    
    row++
next

excel.setColumnWidth(1, 15)
excel.setColumnWidth(3, 15)

if excel.save("demo14_colors.xlsx")
    ? "  Created: demo14_colors.xlsx"
else
    ? "  FAILED: demo14_colors.xlsx"
ok

? ""
? "=============================================="
? "   All demos completed!"
? "=============================================="
? ""
? "Created files:"
? "  1. demo1_simple.xlsx"
? "  2. demo2_multiple_sheets.xlsx"
? "  3. demo3_formulas.xlsx"
? "  4. demo4_styled.xlsx"
? "  5. demo5_dimensions.xlsx"
? "  6. demo6_merged.xlsx"
? "  7. demo7_autofilter.xlsx"
? "  8. demo8_freeze_panes.xlsx"
? "  9. demo9_quick.xlsx"
? "  10. demo10_multi_sheet.xlsx"
? "  11. demo11_images.xlsx"
? "  12. demo12_comprehensive.xlsx"
? "  13. demo13_borders.xlsx"
? "  14. demo14_colors.xlsx"