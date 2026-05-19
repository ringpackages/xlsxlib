load "xlsxlib.ring"

cFileName = substr(filename(),".ring",".xlsx")
? "Generate File: " + cFileName

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

excel.save(cFileName)