load "xlsxlib.ring"

cFileName = substr(filename(),".ring",".xlsx")
? "Generate File: " + cFileName

new ExcelWriter() {
	addSheet("Data")
	setCell(1, 1, "Hello")
	setCell(1, 2, "World")
	setCell(2, 1, 2026)
	save(cFileName)
}