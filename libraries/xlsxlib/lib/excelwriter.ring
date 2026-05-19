/*
    XLSXLib - Excel Library for Ring Programming Language
*/

# ============================================================================
# ExcelWriter Class
# ============================================================================

class ExcelWriter

    # Sheet management
    aSheets
    nCurrentSheet
    
    # Shared strings
    aSharedStrings
    
    # Styles
    aFonts
    aFills
    aBorders
    aCellStyles
    aNumberFormats
    nNextFormatId
    
    # Images
    aImages
    nImageId
    
    # Document properties
    cAuthor
    cTitle
    cCompany
    
    func init
        aSheets = []
        nCurrentSheet = 0
        aSharedStrings = []
        aImages = []
        nImageId = 0
        cAuthor = "RingExcelLib"
        cTitle = "Workbook"
        cCompany = ""
        
        # Initialize default styles
        aFonts = [
            [:name = "Calibri", :size = 11, :bold = false, :italic = false, :color = "000000"]
        ]
        aFills = [
            [:pattern = "none"],
            [:pattern = "gray125"]
        ]
        aBorders = [
            [:left = "", :right = "", :top = "", :bottom = ""]
        ]
        aCellStyles = [
            [:fontId = 0, :fillId = 0, :borderId = 0, :numFmtId = 0, :align = "", :valign = "", :wrap = false]
        ]
        aNumberFormats = []
        nNextFormatId = 164
        
        return self
    
    # ========================================================================
    # Document Properties
    # ========================================================================
    
    func setAuthor author
        cAuthor = author
        return self
    
    func setTitle title
        cTitle = title
        return self
    
    func setCompany company
        cCompany = company
        return self
    
    # ========================================================================
    # Sheet Management
    # ========================================================================
    
    func addSheet name
        sheet = [
            :name = name,
            :cells = [],
            :colWidths = [],
            :rowHeights = [],
            :mergedCells = [],
            :autoFilter = NULL,
            :freezePane = NULL,
            :images = []
        ]
        aSheets + sheet
        nCurrentSheet = len(aSheets)
        return self
    
    func selectSheet index
        if index >= 1 and index <= len(aSheets)
            nCurrentSheet = index
        ok
        return self
    
    func selectSheetByName name
        sheetsLen = len(aSheets)
        for i = 1 to sheetsLen
            if aSheets[i][:name] = name
                nCurrentSheet = i
                exit
            ok
        next
        return self
    
    func getSheetCount
        return len(aSheets)
    
    # ========================================================================
    # Cell Operations
    # ========================================================================
    
    func setCell row, col, value
        if nCurrentSheet = 0 return self ok
        
        cell = [:row = row, :col = col, :value = value, :type = "string", :styleId = 0, :formula = ""]
        
        if isNumber(value)
            cell[:type] = "number"
        elseif isString(value)
            if left(value, 1) = "="
                cell[:type] = "formula"
                cell[:formula] = substr(value, 2)
                cell[:value] = ""
            else
                cell[:type] = "string"
                cell[:value] = addSharedString(value)
            ok
        ok
        
        aSheets[nCurrentSheet][:cells] + cell
        return self
    
    func setCellWithStyle row, col, value, styleId
        if nCurrentSheet = 0 return self ok
        
        cell = [:row = row, :col = col, :value = value, :type = "string", :styleId = styleId, :formula = ""]
        
        if isNumber(value)
            cell[:type] = "number"
        elseif isString(value)
            if left(value, 1) = "="
                cell[:type] = "formula"
                cell[:formula] = substr(value, 2)
                cell[:value] = ""
            else
                cell[:type] = "string"
                cell[:value] = addSharedString(value)
            ok
        ok
        
        aSheets[nCurrentSheet][:cells] + cell
        return self
    
    func setFormula row, col, formula
        return setCell(row, col, "=" + formula)
    
    func setNumber row, col, value
        if nCurrentSheet = 0 return self ok
        
        cell = [:row = row, :col = col, :value = value, :type = "number", :styleId = 0, :formula = ""]
        aSheets[nCurrentSheet][:cells] + cell
        return self
    
    func setDate row, col, year, month, day
        # Excel stores dates as number of days since 1900-01-01
        # This is a simplified calculation
        if nCurrentSheet = 0 return self ok
        
        dateValue = (year - 1900) * 365 + floor((year - 1900) / 4) + 
                    getDaysBeforeMonth(month, year) + day + 1
        
        cell = [:row = row, :col = col, :value = dateValue, :type = "number", :styleId = 0, :formula = ""]
        aSheets[nCurrentSheet][:cells] + cell
        return self
    
    func getDaysBeforeMonth month, year
        days = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
        result = days[month]
        # Add leap day if applicable
        if month > 2 and ((year % 4 = 0 and year % 100 != 0) or year % 400 = 0)
            result++
        ok
        return result
    
    # ========================================================================
    # Bulk Data Operations
    # ========================================================================
    
    func exportList data, startRow, startCol, hasHeader
        if nCurrentSheet = 0 return self ok
        if hasHeader = NULL hasHeader = false ok
        
        row = startRow
        dataLen = len(data)
        for i = 1 to dataLen
            rowData = data[i]
            col = startCol
            rowDataLen = len(rowData)
            for j = 1 to rowDataLen
                if hasHeader and i = 1
                    # Header row - use bold style
                    styleId = createStyle([:bold = true, :bgColor = "D9E1F2"])
                    setCellWithStyle(row, col, rowData[j], styleId)
                else
                    setCell(row, col, rowData[j])
                ok
                col++
            next
            row++
        next
        return self
    
    func setRow row, values, startCol
        if startCol = NULL startCol = 1 ok
        col = startCol
        valuesLen = len(values)
        for i = 1 to valuesLen
            setCell(row, col, values[i])
            col++
        next
        return self
    
    func setColumn col, values, startRow
        if startRow = NULL startRow = 1 ok
        row = startRow
        valuesLen = len(values)
        for i = 1 to valuesLen
            setCell(row, col, values[i])
            row++
        next
        return self
    
    # ========================================================================
    # Styling
    # ========================================================================
    
    func createStyle options
        if options = NULL options = [] ok
        
        # Create font
        fontId = 0
        hasFontChanges = false
        if options[:bold] = true hasFontChanges = true ok
        if options[:italic] = true hasFontChanges = true ok
        if options[:fontSize] != NULL hasFontChanges = true ok
        if options[:fontName] != NULL hasFontChanges = true ok
        if options[:fontColor] != NULL and isString(options[:fontColor]) hasFontChanges = true ok
        
        if hasFontChanges
            font = [:name = "Calibri", :size = 11, :bold = false, :italic = false, :color = "000000"]
            if options[:fontName] != NULL font[:name] = options[:fontName] ok
            if options[:fontSize] != NULL font[:size] = options[:fontSize] ok
            if options[:bold] = true font[:bold] = true ok
            if options[:italic] = true font[:italic] = true ok
            if options[:fontColor] != NULL and isString(options[:fontColor])
                font[:color] = excelColorToHex(options[:fontColor])
            ok
            aFonts + font
            fontId = len(aFonts) - 1
        ok
        
        # Create fill
        fillId = 0
        if options[:bgColor] != NULL and isString(options[:bgColor])
            fill = [:pattern = "solid", :color = excelColorToHex(options[:bgColor])]
            aFills + fill
            fillId = len(aFills) - 1
        ok
        
        # Create border
        borderId = 0
        if options[:border] != NULL or (options[:borderColor] != NULL and isString(options[:borderColor]))
            borderStyle = "thin"
            if options[:border] != NULL and isString(options[:border])
                borderStyle = options[:border]
            ok
            borderColor = "000000"
            if options[:borderColor] != NULL and isString(options[:borderColor])
                borderColor = excelColorToHex(options[:borderColor])
            ok
            border = [:left = borderStyle, :right = borderStyle, :top = borderStyle, :bottom = borderStyle, :color = borderColor]
            aBorders + border
            borderId = len(aBorders) - 1
        ok
        
        # Number format
        numFmtId = 0
        if options[:numberFormat] != NULL
            numFmtId = options[:numberFormat]
        ok
        
        # Alignment
        align = ""
        valign = ""
        wrap = false
        if options[:align] != NULL and isString(options[:align]) align = options[:align] ok
        if options[:valign] != NULL and isString(options[:valign]) valign = options[:valign] ok
        if options[:wrap] = true wrap = true ok
        
        # Create cell style
        style = [:fontId = fontId, :fillId = fillId, :borderId = borderId, 
                 :numFmtId = numFmtId, :align = align, :valign = valign, :wrap = wrap]
        aCellStyles + style
        
        return len(aCellStyles) - 1
    
    func createHeaderStyle
        return createStyle([:bold = true, :bgColor = "4472C4", :fontColor = "FFFFFF", :align = "center"])
    
    func createCurrencyStyle
        return createStyle([:numberFormat = EXCEL_FORMAT_CURRENCY])
    
    func createPercentageStyle
        return createStyle([:numberFormat = EXCEL_FORMAT_PERCENTAGE])
    
    func createDateStyle
        return createStyle([:numberFormat = EXCEL_FORMAT_DATE])
    
    # ========================================================================
    # Column and Row Formatting
    # ========================================================================
    
    func setColumnWidth col, width
        if nCurrentSheet = 0 return self ok
        aSheets[nCurrentSheet][:colWidths] + [col, width]
        return self
    
    func setRowHeight row, height
        if nCurrentSheet = 0 return self ok
        aSheets[nCurrentSheet][:rowHeights] + [row, height]
        return self
    
    func autoFitColumn col
        # Set a reasonable auto-fit width (approximate)
        return setColumnWidth(col, 12)
    
    # ========================================================================
    # Merged Cells
    # ========================================================================
    
    func mergeCells row1, col1, row2, col2
        if nCurrentSheet = 0 return self ok
        aSheets[nCurrentSheet][:mergedCells] + [row1, col1, row2, col2]
        return self
    
    # ========================================================================
    # Auto Filter
    # ========================================================================
    
    func setAutoFilter row1, col1, row2, col2
        if nCurrentSheet = 0 return self ok
        aSheets[nCurrentSheet][:autoFilter] = [row1, col1, row2, col2]
        return self
    
    # ========================================================================
    # Freeze Panes
    # ========================================================================
    
    func freezePanes row, col
        /*
            Freeze rows above and columns to the left of the specified cell
            freezePanes(2, 1) freezes the first row
            freezePanes(1, 2) freezes the first column
            freezePanes(2, 2) freezes first row and first column
        */
        if nCurrentSheet = 0 return self ok
        aSheets[nCurrentSheet][:freezePane] = [row, col]
        return self
    
    func freezeTopRow
        return freezePanes(2, 1)
    
    func freezeFirstColumn
        return freezePanes(1, 2)
    
    # ========================================================================
    # Images
    # ========================================================================
    
    func addImage imagePath, row, col, width, height
        /*
            Add an image to the current sheet
            imagePath: path to image file (png, jpg, bmp, gif)
            row, col: cell position for the image
            width: width in cells (default 5)
            height: height in cells (default 5)
        */
        if nCurrentSheet = 0 return self ok
        if width = NULL width = 5 ok
        if height = NULL height = 5 ok
        
        if !fexists(imagePath)
            ? "Warning: Image file not found: " + imagePath
            return self
        ok
        
        imageData = read(imagePath)
        if len(imageData) = 0
            ? "Warning: Could not read image file: " + imagePath
            return self
        ok
        
        ext = lower(excelGetImageExtension(imagePath))
        contentType = "image/png"
        
        switch ext
            on "png"
                contentType = "image/png"
            on "jpg"
                contentType = "image/jpeg"
            on "jpeg"
                contentType = "image/jpeg"
            on "gif"
                contentType = "image/gif"
            on "bmp"
                contentType = "image/bmp"
        off
        
        nImageId++
        
        img = [
            :id = nImageId,
            :filename = "image" + nImageId + "." + ext,
            :data = imageData,
            :contentType = contentType,
            :row = row,
            :col = col,
            :width = width,
            :height = height,
            :sheet = nCurrentSheet
        ]
        
        aImages + img
        aSheets[nCurrentSheet][:images] + nImageId
        
        return self
    
    func addImageCentered imagePath, row, col, width, height
        return addImage(imagePath, row, col, width, height)
          
    # ========================================================================
    # Save Document
    # ========================================================================
    
    func save filename
        sep = excelGetSep()
        
        # Create temp directory structure
        tempDir = filename + "_temp" + sep
        excelMakeDir(tempDir)
        excelMakeDir(tempDir + "_rels")
        excelMakeDir(tempDir + "docProps")
        excelMakeDir(tempDir + "xl")
        excelMakeDir(tempDir + "xl" + sep + "_rels")
        excelMakeDir(tempDir + "xl" + sep + "worksheets")
        
        # Create media folder if we have images
        if len(aImages) > 0
            excelMakeDir(tempDir + "xl" + sep + "media")
            excelMakeDir(tempDir + "xl" + sep + "drawings")
            excelMakeDir(tempDir + "xl" + sep + "drawings" + sep + "_rels")
        ok
        
        # Write all XML files
        writeContentTypes(tempDir)
        writeRels(tempDir)
        writeWorkbookRels(tempDir)
        writeWorkbook(tempDir)
        writeSharedStrings(tempDir)
        writeStyles(tempDir)
        writeCore(tempDir)
        writeApp(tempDir)
        
        # Write worksheets
        sheetsLen = len(aSheets)
        for i = 1 to sheetsLen
            writeWorksheet(tempDir, i)
        next
        
        # Write drawings if we have images
        if len(aImages) > 0
            writeDrawings(tempDir)
        ok
        
        # Create ZIP file
        result = createZip(tempDir, filename)
        
        # Clean up temp directory
        if isWindows()
            system('rmdir /s /q "' + tempDir + '" 2>nul')
        else
            system("rm -rf '" + tempDir + "'")
        ok
        
        return result
    
    # ========================================================================
    # Helper: Add Shared String
    # ========================================================================
    
    func addSharedString str
        str = "" + str
        stringsLen = len(aSharedStrings)
        for i = 1 to stringsLen
            if aSharedStrings[i] = str
                return i - 1
            ok
        next
        aSharedStrings + str
        return len(aSharedStrings) - 1
    
    # ========================================================================
    # XML Generation Methods
    # ========================================================================
    
    func writeContentTypes tempDir
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        c += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        c += '<Default Extension="xml" ContentType="application/xml"/>'
        
        # Image types
        imagesLen = len(aImages)
        hasPng = false
        hasJpg = false
        hasGif = false
        hasBmp = false
        
        for i = 1 to imagesLen
            ct = aImages[i][:contentType]
            if ct = "image/png" and !hasPng
                c += '<Default Extension="png" ContentType="image/png"/>'
                hasPng = true
            ok
            if ct = "image/jpeg" and !hasJpg
                c += '<Default Extension="jpg" ContentType="image/jpeg"/>'
                c += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                hasJpg = true
            ok
            if ct = "image/gif" and !hasGif
                c += '<Default Extension="gif" ContentType="image/gif"/>'
                hasGif = true
            ok
            if ct = "image/bmp" and !hasBmp
                c += '<Default Extension="bmp" ContentType="image/bmp"/>'
                hasBmp = true
            ok
        next
        
        c += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        c += '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        c += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        c += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        c += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        
        sheetsLen = len(aSheets)
        for i = 1 to sheetsLen
            c += '<Override PartName="/xl/worksheets/sheet' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            
            # Drawing for this sheet if it has images
            sheetImagesLen = len(aSheets[i][:images])
            if sheetImagesLen > 0
                c += '<Override PartName="/xl/drawings/drawing' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            ok
        next
        
        c += '</Types>'
        
        write(tempDir + "[Content_Types].xml", c)
    
    func writeRels tempDir
        sep = excelGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        c += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        c += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        c += '</Relationships>'
        
        write(tempDir + "_rels" + sep + ".rels", c)
    
    func writeWorkbookRels tempDir
        sep = excelGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        
        relId = 1
        sheetsLen = len(aSheets)
        for i = 1 to sheetsLen
            c += '<Relationship Id="rId' + relId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + i + '.xml"/>'
            relId++
        next
        
        c += '<Relationship Id="rId' + relId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        relId++
        c += '<Relationship Id="rId' + relId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        
        c += '</Relationships>'
        
        write(tempDir + "xl" + sep + "_rels" + sep + "workbook.xml.rels", c)
    
    func writeWorkbook tempDir
        sep = excelGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        c += '<sheets>'
        
        sheetsLen = len(aSheets)
        for i = 1 to sheetsLen
            c += '<sheet name="' + excelXmlEsc(aSheets[i][:name]) + '" sheetId="' + i + '" r:id="rId' + i + '"/>'
        next
        
        c += '</sheets>'
        c += '</workbook>'
        
        write(tempDir + "xl" + sep + "workbook.xml", c)
    
    func writeSharedStrings tempDir
        sep = excelGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        stringsLen = len(aSharedStrings)
        c += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + stringsLen + '" uniqueCount="' + stringsLen + '">'
        
        for i = 1 to stringsLen
            c += '<si><t>' + excelXmlEsc(aSharedStrings[i]) + '</t></si>'
        next
        
        c += '</sst>'
        
        write(tempDir + "xl" + sep + "sharedStrings.xml", c)
    
    func writeStyles tempDir
        sep = excelGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        
        # Number formats
        numFmtsLen = len(aNumberFormats)
        if numFmtsLen > 0
            c += '<numFmts count="' + numFmtsLen + '">'
            for i = 1 to numFmtsLen
                c += '<numFmt numFmtId="' + aNumberFormats[i][:id] + '" formatCode="' + excelXmlEsc(aNumberFormats[i][:code]) + '"/>'
            next
            c += '</numFmts>'
        ok
        
        # Fonts
        fontsLen = len(aFonts)
        c += '<fonts count="' + fontsLen + '">'
        for i = 1 to fontsLen
            f = aFonts[i]
            c += '<font>'
            if f[:bold] c += '<b/>' ok
            if f[:italic] c += '<i/>' ok
            c += '<sz val="' + f[:size] + '"/>'
            c += '<color rgb="FF' + f[:color] + '"/>'
            c += '<name val="' + f[:name] + '"/>'
            c += '</font>'
        next
        c += '</fonts>'
        
        # Fills
        fillsLen = len(aFills)
        c += '<fills count="' + fillsLen + '">'
        for i = 1 to fillsLen
            f = aFills[i]
            c += '<fill>'
            if f[:pattern] = "solid"
                c += '<patternFill patternType="solid">'
                c += '<fgColor rgb="FF' + f[:color] + '"/>'
                c += '<bgColor indexed="64"/>'
                c += '</patternFill>'
            elseif f[:pattern] = "gray125"
                c += '<patternFill patternType="gray125"/>'
            else
                c += '<patternFill patternType="none"/>'
            ok
            c += '</fill>'
        next
        c += '</fills>'
        
        # Borders
        bordersLen = len(aBorders)
        c += '<borders count="' + bordersLen + '">'
        for i = 1 to bordersLen
            b = aBorders[i]
            c += '<border>'
            if len(b[:left]) > 0
                c += '<left style="' + b[:left] + '"><color auto="1"/></left>'
            else
                c += '<left/>'
            ok
            if len(b[:right]) > 0
                c += '<right style="' + b[:right] + '"><color auto="1"/></right>'
            else
                c += '<right/>'
            ok
            if len(b[:top]) > 0
                c += '<top style="' + b[:top] + '"><color auto="1"/></top>'
            else
                c += '<top/>'
            ok
            if len(b[:bottom]) > 0
                c += '<bottom style="' + b[:bottom] + '"><color auto="1"/></bottom>'
            else
                c += '<bottom/>'
            ok
            c += '<diagonal/>'
            c += '</border>'
        next
        c += '</borders>'
        
        # Cell style XFs (formatting records)
        c += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        
        # Cell XFs
        stylesLen = len(aCellStyles)
        c += '<cellXfs count="' + stylesLen + '">'
        for i = 1 to stylesLen
            s = aCellStyles[i]
            c += '<xf numFmtId="' + s[:numFmtId] + '" fontId="' + s[:fontId] + '" fillId="' + s[:fillId] + '" borderId="' + s[:borderId] + '"'
            
            if s[:fontId] > 0 c += ' applyFont="1"' ok
            if s[:fillId] > 0 c += ' applyFill="1"' ok
            if s[:borderId] > 0 c += ' applyBorder="1"' ok
            if s[:numFmtId] > 0 c += ' applyNumberFormat="1"' ok
            
            if len(s[:align]) > 0 or len(s[:valign]) > 0 or s[:wrap]
                c += ' applyAlignment="1">'
                c += '<alignment'
                if len(s[:align]) > 0 c += ' horizontal="' + s[:align] + '"' ok
                if len(s[:valign]) > 0 c += ' vertical="' + s[:valign] + '"' ok
                if s[:wrap] c += ' wrapText="1"' ok
                c += '/>'
                c += '</xf>'
            else
                c += '/>'
            ok
        next
        c += '</cellXfs>'
        
        c += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        c += '</styleSheet>'
        
        write(tempDir + "xl" + sep + "styles.xml", c)
    
    func writeCore tempDir
        sep = excelGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        c += 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        c += 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        c += '<dc:title>' + excelXmlEsc(cTitle) + '</dc:title>'
        c += '<dc:creator>' + excelXmlEsc(cAuthor) + '</dc:creator>'
        c += '<cp:lastModifiedBy>' + excelXmlEsc(cAuthor) + '</cp:lastModifiedBy>'
        c += '<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>'
        c += '<dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>'
        c += '</cp:coreProperties>'
        
        write(tempDir + "docProps" + sep + "core.xml", c)
    
    func writeApp tempDir
        sep = excelGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        c += '<Application>RingExcelLib</Application>'
        c += '<AppVersion>2.0</AppVersion>'
        if len(cCompany) > 0
            c += '<Company>' + excelXmlEsc(cCompany) + '</Company>'
        ok
        c += '</Properties>'
        
        write(tempDir + "docProps" + sep + "app.xml", c)
    
    func writeWorksheet tempDir, sheetIndex
        sep = excelGetSep()
        sheet = aSheets[sheetIndex]
        cells = sheet[:cells]
        colWidths = sheet[:colWidths]
        rowHeights = sheet[:rowHeights]
        mergedCells = sheet[:mergedCells]
        filterData = sheet[:autoFilter]
        freezePane = sheet[:freezePane]
        sheetImages = sheet[:images]
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        
        # Calculate dimensions
        if len(cells) > 0
            maxRow = 1
            maxCol = 1
            cellsLen = len(cells)
            for i = 1 to cellsLen
                if cells[i][:row] > maxRow maxRow = cells[i][:row] ok
                if cells[i][:col] > maxCol maxCol = cells[i][:col] ok
            next
            c += '<dimension ref="A1:' + excelColLetter(maxCol) + maxRow + '"/>'
        ok
        
        # Sheet views (freeze panes)
        c += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"'
        if freezePane != NULL
            c += '>'
            fr = freezePane[1]
            fc = freezePane[2]
            topLeft = excelColLetter(fc) + fr
            c += '<pane'
            if fc > 1 c += ' xSplit="' + (fc - 1) + '"' ok
            if fr > 1 c += ' ySplit="' + (fr - 1) + '"' ok
            c += ' topLeftCell="' + topLeft + '" activePane="bottomRight" state="frozen"/>'
            c += '</sheetView>'
        else
            c += '/>'
        ok
        c += '</sheetViews>'
        
        c += '<sheetFormatPr defaultRowHeight="15"/>'
        
        # Column widths
        colWidthsLen = len(colWidths)
        if colWidthsLen > 0
            c += '<cols>'
            for i = 1 to colWidthsLen
                colNum = colWidths[i][1]
                colW = colWidths[i][2]
                c += '<col min="' + colNum + '" max="' + colNum + '" width="' + colW + '" customWidth="1"/>'
            next
            c += '</cols>'
        ok
        
        # Sheet data
        c += '<sheetData>'
        
        cellsLen = len(cells)
        if cellsLen > 0
            # Group cells by row
            rowNums = []
            for i = 1 to cellsLen
                rn = cells[i][:row]
                found = false
                rowNumsLen = len(rowNums)
                for j = 1 to rowNumsLen
                    if rowNums[j] = rn
                        found = true
                        exit
                    ok
                next
                if !found
                    rowNums + rn
                ok
            next
            
            # Sort row numbers
            rowNumsLen = len(rowNums)
            for x = 1 to rowNumsLen - 1
                for y = x + 1 to rowNumsLen
                    if rowNums[y] < rowNums[x]
                        temp = rowNums[x]
                        rowNums[x] = rowNums[y]
                        rowNums[y] = temp
                    ok
                next
            next
            
            # Output rows
            for ri = 1 to rowNumsLen
                rn = rowNums[ri]
                rowCells = []
                
                for i = 1 to cellsLen
                    if cells[i][:row] = rn
                        rowCells + cells[i]
                    ok
                next
                
                # Sort cells by column
                rowCellsLen = len(rowCells)
                for x = 1 to rowCellsLen - 1
                    rcLen = len(rowCells)
                    for y = x + 1 to rcLen
                        if rowCells[y][:col] < rowCells[x][:col]
                            temp = rowCells[x]
                            rowCells[x] = rowCells[y]
                            rowCells[y] = temp
                        ok
                    next
                next
                
                # Check for custom row height
                rowAttr = ""
                rowHeightsLen = len(rowHeights)
                for i = 1 to rowHeightsLen
                    if rowHeights[i][1] = rn
                        rowAttr = ' ht="' + rowHeights[i][2] + '" customHeight="1"'
                        exit
                    ok
                next
                
                c += '<row r="' + rn + '"' + rowAttr + '>'
                rowCellsLen = len(rowCells)
                for i = 1 to rowCellsLen
                    cell = rowCells[i]
                    cellCol = cell[:col]
                    cellRow = cell[:row]
                    cellRef = excelColLetter(cellCol) + "" + cellRow
                    
                    c += '<c r="' + cellRef + '"'
                    cellStyleId = cell[:styleId]
                    if isNumber(cellStyleId) and cellStyleId > 0
                        c += ' s="' + cellStyleId + '"'
                    ok
                    
                    cellValue = cell[:value]
                    cellType = cell[:type]
                    
                    # Convert value to string safely
                    if isNumber(cellValue)
                        cellValueStr = "" + cellValue
                    elseif isString(cellValue)
                        cellValueStr = cellValue
                    else
                        cellValueStr = ""
                    ok
                    
                    if cellType = "string"
                        c += ' t="s"><v>' + cellValueStr + '</v></c>'
                    elseif cellType = "number"
                        c += '><v>' + cellValueStr + '</v></c>'
                    elseif cellType = "formula"
                        cellFormula = cell[:formula]
                        if isString(cellFormula)
                            c += '><f>' + excelXmlEsc(cellFormula) + '</f></c>'
                        else
                            c += '/>'
                        ok
                    else
                        c += '/>'
                    ok
                next
                c += '</row>'
            next
        ok
        
        c += '</sheetData>'
        
        # Auto filter (must come before mergeCells per ECMA-376)
        if isList(filterData) and len(filterData) = 4
            c += '<autoFilter ref="' + excelColLetter(filterData[2]) + filterData[1] + ':' + excelColLetter(filterData[4]) + filterData[3] + '"/>'
        ok
        
        # Merged cells (must come after autoFilter)
        mergedLen = len(mergedCells)
        if mergedLen > 0
            c += '<mergeCells count="' + mergedLen + '">'
            for i = 1 to mergedLen
                m = mergedCells[i]
                c += '<mergeCell ref="' + excelColLetter(m[2]) + m[1] + ':' + excelColLetter(m[4]) + m[3] + '"/>'
            next
            c += '</mergeCells>'
        ok
        
        # Page margins MUST come before drawing
        c += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        
        # Drawing reference if sheet has images (MUST come after pageMargins)
        sheetImagesLen = len(sheetImages)
        if sheetImagesLen > 0
            c += '<drawing r:id="rId1"/>'
        ok
        
        c += '</worksheet>'
        
        write(tempDir + "xl" + sep + "worksheets" + sep + "sheet" + sheetIndex + ".xml", c)
        
        # Write worksheet relationships if there are images
        if sheetImagesLen > 0
            writeWorksheetRels(tempDir, sheetIndex)
        ok
    
    func writeWorksheetRels tempDir, sheetIndex
        sep = excelGetSep()
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' + sheetIndex + '.xml"/>'
        c += '</Relationships>'
        
        # Create worksheets/_rels folder if needed
        excelMakeDir(tempDir + "xl" + sep + "worksheets" + sep + "_rels")
        write(tempDir + "xl" + sep + "worksheets" + sep + "_rels" + sep + "sheet" + sheetIndex + ".xml.rels", c)
    
    func writeDrawings tempDir
        sep = excelGetSep()
        
        sheetsLen = len(aSheets)
        for sheetIdx = 1 to sheetsLen
            sheetImages = aSheets[sheetIdx][:images]
            sheetImagesLen = len(sheetImages)
            
            if sheetImagesLen > 0
                # Write drawing XML
                c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                c += '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                c += 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                
                relId = 1
                for i = 1 to sheetImagesLen
                    imgId = sheetImages[i]
                    
                    # Find image info
                    img = NULL
                    imagesLen = len(aImages)
                    for j = 1 to imagesLen
                        if aImages[j][:id] = imgId
                            img = aImages[j]
                            exit
                        ok
                    next
                    
                    if img != NULL
                        c += '<xdr:twoCellAnchor>'
                        c += '<xdr:from><xdr:col>' + (img[:col] - 1) + '</xdr:col><xdr:colOff>0</xdr:colOff>'
                        c += '<xdr:row>' + (img[:row] - 1) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                        c += '<xdr:to><xdr:col>' + (img[:col] + img[:width] - 1) + '</xdr:col><xdr:colOff>0</xdr:colOff>'
                        c += '<xdr:row>' + (img[:row] + img[:height] - 1) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                        c += '<xdr:pic>'
                        c += '<xdr:nvPicPr>'
                        c += '<xdr:cNvPr id="' + img[:id] + '" name="Picture ' + img[:id] + '"/>'
                        c += '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>'
                        c += '</xdr:nvPicPr>'
                        c += '<xdr:blipFill>'
                        c += '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId' + relId + '"/>'
                        c += '<a:stretch><a:fillRect/></a:stretch>'
                        c += '</xdr:blipFill>'
                        c += '<xdr:spPr>'
                        c += '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
                        c += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                        c += '</xdr:spPr>'
                        c += '</xdr:pic>'
                        c += '<xdr:clientData/>'
                        c += '</xdr:twoCellAnchor>'
                        
                        relId++
                    ok
                next
                
                c += '</xdr:wsDr>'
                
                write(tempDir + "xl" + sep + "drawings" + sep + "drawing" + sheetIdx + ".xml", c)
                
                # Write drawing relationships
                writeDrawingRels(tempDir, sheetIdx)
            ok
        next
    
    func writeDrawingRels tempDir, sheetIdx
        sep = excelGetSep()
        sheetImages = aSheets[sheetIdx][:images]
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        
        relId = 1
        sheetImagesLen = len(sheetImages)
        for i = 1 to sheetImagesLen
            imgId = sheetImages[i]
            
            # Find image info
            imagesLen = len(aImages)
            for j = 1 to imagesLen
                if aImages[j][:id] = imgId
                    c += '<Relationship Id="rId' + relId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' + aImages[j][:filename] + '"/>'
                    relId++
                    exit
                ok
            next
        next
        
        c += '</Relationships>'
        
        write(tempDir + "xl" + sep + "drawings" + sep + "_rels" + sep + "drawing" + sheetIdx + ".xml.rels", c)
    
    func createZip tempDir, filename
        sep = excelGetSep()
        
        filesList = []
        
        filesList + ["[Content_Types].xml", read(tempDir + "[Content_Types].xml")]
        filesList + ["_rels/.rels", read(tempDir + "_rels" + sep + ".rels")]
        filesList + ["docProps/core.xml", read(tempDir + "docProps" + sep + "core.xml")]
        filesList + ["docProps/app.xml", read(tempDir + "docProps" + sep + "app.xml")]
        filesList + ["xl/workbook.xml", read(tempDir + "xl" + sep + "workbook.xml")]
        filesList + ["xl/_rels/workbook.xml.rels", read(tempDir + "xl" + sep + "_rels" + sep + "workbook.xml.rels")]
        filesList + ["xl/styles.xml", read(tempDir + "xl" + sep + "styles.xml")]
        filesList + ["xl/sharedStrings.xml", read(tempDir + "xl" + sep + "sharedStrings.xml")]
        
        sheetsLen = len(aSheets)
        for i = 1 to sheetsLen
            sheetPath = "xl/worksheets/sheet" + i + ".xml"
            localPath = tempDir + "xl" + sep + "worksheets" + sep + "sheet" + i + ".xml"
            filesList + [sheetPath, read(localPath)]
            
            # Worksheet relationships if exists
            relsPath = tempDir + "xl" + sep + "worksheets" + sep + "_rels" + sep + "sheet" + i + ".xml.rels"
            if fexists(relsPath)
                filesList + ["xl/worksheets/_rels/sheet" + i + ".xml.rels", read(relsPath)]
            ok
        next
        
        # Add drawings
        for i = 1 to sheetsLen
            drawingPath = tempDir + "xl" + sep + "drawings" + sep + "drawing" + i + ".xml"
            if fexists(drawingPath)
                filesList + ["xl/drawings/drawing" + i + ".xml", read(drawingPath)]
            ok
            
            drawingRelsPath = tempDir + "xl" + sep + "drawings" + sep + "_rels" + sep + "drawing" + i + ".xml.rels"
            if fexists(drawingRelsPath)
                filesList + ["xl/drawings/_rels/drawing" + i + ".xml.rels", read(drawingRelsPath)]
            ok
        next
        
        # Add images
        imagesLen = len(aImages)
        for i = 1 to imagesLen
            filesList + ["xl/media/" + aImages[i][:filename], aImages[i][:data]]
        next
        
        return excelZipCreateFile(filename, filesList)
