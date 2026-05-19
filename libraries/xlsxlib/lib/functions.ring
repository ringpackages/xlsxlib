/*
    XLSXLib - Excel Library for Ring Programming Language
*/

# ============================================================================
# Quick Helper Functions
# ============================================================================

func listToExcel data, filename, sheetName, hasHeader
    if sheetName = NULL sheetName = "Sheet1" ok
    if hasHeader = NULL hasHeader = false ok
    
    excel = new ExcelWriter()
    excel.addSheet(sheetName)
    excel.exportList(data, 1, 1, hasHeader)
    return excel.save(filename)

func listsToExcel dataList, filename
    excel = new ExcelWriter()
    
    listLen = len(dataList)
    for i = 1 to listLen
        entry = dataList[i]
        sheetName = entry[1]
        data = entry[2]
        hasHeader = false
        entryLen = len(entry)
        if entryLen > 2
            hasHeader = entry[3]
        ok
        
        excel.addSheet(sheetName)
        excel.exportList(data, 1, 1, hasHeader)
    next
    
    return excel.save(filename)

func quickExcel filename, data, sheetName
    /*
        Quick function to create a simple Excel file
    */
    if sheetName = NULL sheetName = "Sheet1" ok
    return listToExcel(data, filename, sheetName, true)

# ============================================================================
# Platform Helper Functions
# ============================================================================

func excelGetSep
    if isWindows()
        return "\"
    else
        return "/"
    ok

func excelXmlEsc str
    str = "" + str
    str = substr(str, "&", "&amp;")
    str = substr(str, "<", "&lt;")
    str = substr(str, ">", "&gt;")
    str = substr(str, '"', "&quot;")
    return str

func excelColLetter col
    result = ""
    while col > 0
        col = col - 1
        remainder = col % 26
        result = char(65 + remainder) + result
        col = floor(col / 26)
    end
    return result

func excelMakeDir path
    if isWindows()
        path = substr(path, "/", "\")
        system('mkdir "' + path + '" 2>nul')
    else
        system("mkdir -p '" + path + "'")
    ok

func excelGetImageExtension filepath
    # Find the last dot in the filepath
    dotPos = 0
    fpLen = len(filepath)
    for i = 1 to fpLen
        if substr(filepath, i, 1) = "."
            dotPos = i
        ok
    next
    if dotPos > 0 and dotPos < fpLen
        return substr(filepath, dotPos + 1, fpLen - dotPos)
    ok
    return "png"

func excelColorToHex color
    # Convert color name to hex
    # Safety check
    if color = NULL return "000000" ok
    if !isString(color) 
        color = "" + color
    ok
    
    color = lower(color)
    colors = [
        :black = "000000",
        :white = "FFFFFF",
        :red = "FF0000",
        :green = "00FF00",
        :blue = "0000FF",
        :yellow = "FFFF00",
        :orange = "FFA500",
        :purple = "800080",
        :gray = "808080",
        :grey = "808080",
        :navy = "000080",
        :teal = "008080",
        :maroon = "800000",
        :silver = "C0C0C0",
        :lime = "00FF00",
        :aqua = "00FFFF",
        :fuchsia = "FF00FF",
        :olive = "808000"
    ]
    
    if colors[color] != NULL
        return colors[color]
    ok
    
    # Remove # if present
    if left(color, 1) = "#"
        color = substr(color, 2)
    ok
    
    return upper(color)

# ============================================================================
# ZIP Functions for Excel-Compatible ZIP Creation
# ============================================================================

func excelZipInitCRC
    if len(aExcelZipCRC32Table) > 0
        return
    ok
    
    for i = 0 to 255
        crc = i
        for j = 1 to 8
            if (crc & 1) = 1
                crc = (crc >> 1) ^ 0xEDB88320
            else
                crc = crc >> 1
            ok
        next
        aExcelZipCRC32Table + crc
    next

func excelZipCRC32 data
    excelZipInitCRC()
    
    crc = 0xFFFFFFFF
    dataLen = len(data)
    for i = 1 to dataLen
        b = ascii(data[i])
        idx = ((crc ^ b) & 0xFF) + 1
        crc = (crc >> 8) ^ aExcelZipCRC32Table[idx]
    next
    
    return crc ^ 0xFFFFFFFF

func excelZipWord value
    return char(value & 0xFF) + char((value >> 8) & 0xFF)

func excelZipDWord value
    return char(value & 0xFF) + char((value >> 8) & 0xFF) + char((value >> 16) & 0xFF) + char((value >> 24) & 0xFF)

func excelZipCreateFile filename, filesList
    output = ""
    centralDir = ""
    offset = 0
    
    filesCount = len(filesList)
    for i = 1 to filesCount
        entry = filesList[i]
        zipPath = entry[1]
        content = entry[2]
        
        crc32 = excelZipCRC32(content)
        uncompSize = len(content)
        compMethod = 0
        compSize = uncompSize
        fileData = content
        
        dosTime = 0x0000
        dosDate = 0x0021
        
        localHeader = ""
        localHeader += char(0x50) + char(0x4B) + char(0x03) + char(0x04)
        localHeader += excelZipWord(20)
        localHeader += excelZipWord(0)
        localHeader += excelZipWord(compMethod)
        localHeader += excelZipWord(dosTime)
        localHeader += excelZipWord(dosDate)
        localHeader += excelZipDWord(crc32)
        localHeader += excelZipDWord(compSize)
        localHeader += excelZipDWord(uncompSize)
        localHeader += excelZipWord(len(zipPath))
        localHeader += excelZipWord(0)
        localHeader += zipPath
        
        output += localHeader
        output += fileData
        
        centralEntry = ""
        centralEntry += char(0x50) + char(0x4B) + char(0x01) + char(0x02)
        centralEntry += excelZipWord(20)
        centralEntry += excelZipWord(20)
        centralEntry += excelZipWord(0)
        centralEntry += excelZipWord(compMethod)
        centralEntry += excelZipWord(dosTime)
        centralEntry += excelZipWord(dosDate)
        centralEntry += excelZipDWord(crc32)
        centralEntry += excelZipDWord(compSize)
        centralEntry += excelZipDWord(uncompSize)
        centralEntry += excelZipWord(len(zipPath))
        centralEntry += excelZipWord(0)
        centralEntry += excelZipWord(0)
        centralEntry += excelZipWord(0)
        centralEntry += excelZipWord(0)
        centralEntry += excelZipDWord(0)
        centralEntry += excelZipDWord(offset)
        centralEntry += zipPath
        
        centralDir += centralEntry
        offset += len(localHeader) + len(fileData)
    next
    
    eocd = ""
    eocd += char(0x50) + char(0x4B) + char(0x05) + char(0x06)
    eocd += excelZipWord(0)
    eocd += excelZipWord(0)
    eocd += excelZipWord(filesCount)
    eocd += excelZipWord(filesCount)
    eocd += excelZipDWord(len(centralDir))
    eocd += excelZipDWord(offset)
    eocd += excelZipWord(0)
    
    output += centralDir
    output += eocd
    
    write(filename, output)
    return fexists(filename)

