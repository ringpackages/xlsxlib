load "xlsxlib.ring"

cFileName = substr(filename(),".ring",".xlsx")
? "Generate File: " + cFileName

data = [
    ["Name", "Age", "City"],
    ["Ahmed", 25, "Cairo"],
    ["Rola", 30, "Riyadh"]
]

quickExcel(cFileName, data, "People")