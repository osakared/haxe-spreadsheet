package spreadsheet.format;

import haxe.io.Bytes;
import haxe.io.Output;
import haxe.Resource;
import haxe.zip.Entry;

typedef SharedStrings = Map<String, Array<Xml>>;

class XlsxWriter
{
    static public function write(spreadsheet:Spreadsheet, output:Output)
    {
        var zipWriter = new haxe.zip.Writer(output);

        var entries = new List<Entry>();

        var prefabFiles = [
            '[Content_Types].xml',
            '_rels/.rels',
            'docProps/app.xml',
            'docProps/core.xml',
            'xl/styles.xml',
            'xl/theme/theme1.xml'
        ];

        // Because of "shared strings" we gotta go through once to assemble these strings, then print to string after finalizing
        var sharedStrings = new SharedStrings();
        var sheetXmls = new Array<Xml>();

        for (sheetIndex in 0...spreadsheet.numSheets()) {
            var sheet = spreadsheet.sheetAt(sheetIndex);
            var sheetXml = xmlFromSheet(sheet, sharedStrings);
            sheetXmls.push(sheetXml);
        }

        // Make shared strings
        var sharedStringsXml = xmlFromSharedStrings(sharedStrings);
        var sharedStringsBytes = Bytes.ofString(sharedStringsXml.toString());
        entries.push(bytesToEntry(sharedStringsBytes, 'xl/sharedStrings.xml'));

        // Save the sheet xmls
        var sheetIndex = 0;
        for (sheetXml in sheetXmls) {
            var sheetBytes = Bytes.ofString(sheetXml.toString());
            entries.push(bytesToEntry(sheetBytes, 'xl/worksheets/sheet${sheetIndex + 1}.xml'));
            sheetIndex++;
        }

        // Save files that reference the different sheets
        saveSheetRefXmls(spreadsheet, entries);

        for (prefabFile in prefabFiles) {
            var bytes = Resource.getBytes('xlsx/' + prefabFile);
            entries.push(bytesToEntry(bytes, prefabFile));
        }

        zipWriter.write(entries);
    }

    static private function saveSheetRefXmls(spreadsheet:Spreadsheet, entries:List<Entry>)
    {
        var workbookRels = Xml.createElement('Relationships');
        workbookRels.set('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        var workbook = Xml.createElement('workbook');
        workbook.set('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        workbook.set('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        workbook.addChild(Xml.createElement('workbookPr'));
        workbook.addChild(Xml.createElement('workbookProtection'));

        // Really wish haxe already had them inline xml literals
        var bookViews = Xml.createElement('bookViews');
        var workbookView = Xml.createElement('workbookView');
        workbookView.set('showHorizontalScroll', 'true');
        workbookView.set('showVerticalScroll', 'true');
        workbookView.set('showSheetTabs', 'true');
        workbookView.set('xWindow', '0');
        workbookView.set('yWindow', '0');
        workbookView.set('windowWidth', '16384');
        workbookView.set('windowHeight', '8192');
        workbookView.set('tabRatio', '500');
        workbookView.set('firstSheet', '0');
        workbookView.set('activeTab', '0');
        bookViews.addChild(workbookView);
        workbook.addChild(bookViews);

        var calcPr = Xml.createElement('calcPr');
        calcPr.set('iterateCount', '100');
        calcPr.set('refMode', 'A1');
        calcPr.set('iterate', 'false');
        calcPr.set('iterateDelta', '0.001');
        workbook.addChild(calcPr);

        var sheets = Xml.createElement('sheets');

        var refID = 1;

        for (sheetIndex in 0...spreadsheet.numSheets()) {
            var sheet = spreadsheet.sheetAt(sheetIndex);
            var relationship = Xml.createElement('Relationship');
            relationship.set('Id', 'rId$refID');
            relationship.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
            relationship.set('Target', 'worksheets/sheet$refID.xml');
            workbookRels.addChild(relationship);

            var sheetXml = Xml.createElement('sheet');
            sheetXml.set('name', sheet.name);
            sheetXml.set('sheetId', Std.string(refID));
            sheetXml.set('state', 'visible');
            sheetXml.set('r:id', 'rId$refID');
            sheets.addChild(sheetXml);

            refID++;
        }

        workbook.addChild(sheets);

        var styleRelationship = Xml.createElement('Relationship');
        styleRelationship.set('Id', 'rId$refID');
        styleRelationship.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');
        styleRelationship.set('Target', 'worksheets/sheet$refID.xml');
        workbookRels.addChild(styleRelationship);
        refID++;

        var sharedStringsRelationship = Xml.createElement('Relationship');
        sharedStringsRelationship.set('Id', 'rId$refID');
        sharedStringsRelationship.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings');
        sharedStringsRelationship.set('Target', 'sharedStrings.xml');
        workbookRels.addChild(sharedStringsRelationship);

        var workbookRelsBytes = Bytes.ofString(workbookRels.toString());
        entries.push(bytesToEntry(workbookRelsBytes, 'xl/_rels/workbook.xml.rels'));

        var workbookBytes = Bytes.ofString(workbook.toString());
        entries.push(bytesToEntry(workbookBytes, 'xl/workbook.xml'));
    }

    static private function bytesToEntry(bytes:Bytes, filename:String):Entry
    {
        return {
            fileName: filename,
            fileSize: bytes.length,
            fileTime: Date.now(),
            compressed: false,
            dataSize: bytes.length,
            data: bytes,
            crc32: haxe.crypto.Crc32.make(bytes),
            extraFields: null
        }
    }

    static private function xmlFromSharedStrings(sharedStrings:SharedStrings):Xml
    {
        var sharedStringsXml = Xml.createElement('sst');
        sharedStringsXml.set('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        var uniqueCount:Int = 0;
        for (sharedString in sharedStrings.keys()) {
            var si = Xml.createElement('si');
            var t = Xml.createElement('t');
            var pcData = Xml.createPCData(sharedString);
            t.addChild(pcData);
            si.addChild(t);
            sharedStringsXml.addChild(si);

            var likeElements = sharedStrings[sharedString];
            for (likeElement in likeElements) {
                var referenceData = Xml.createPCData(Std.string(uniqueCount));
                likeElement.addChild(referenceData);
            }

            uniqueCount++;
        }
        sharedStringsXml.set('uniqueCount', Std.string(uniqueCount));

        return sharedStringsXml;
    }

    static private function xmlFromSheet(sheet:Sheet, sharedStrings:SharedStrings):Xml
    {
        // TODO rewrite all of this to be easier to read with inline xml once haxe adds support for that
        var worksheet = Xml.createElement('worksheet');
        worksheet.set('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // sheetPr
        var sheetPr = Xml.createElement('sheetPr');
        var outlinePr = Xml.createElement('outlinePr');
        outlinePr.set('summaryBelow', '1');
        outlinePr.set('summaryRight', '1');
        sheetPr.addChild(outlinePr);
        sheetPr.addChild(Xml.createElement('pageSetUpPr'));
        worksheet.addChild(sheetPr);

        // sheetViews
        var sheetViews = Xml.createElement('sheetViews');
        var sheetView = Xml.createElement('sheetView');
        var selection = Xml.createElement('selection');
        selection.set('activeCell', 'A1');
        selection.set('sqref', 'A1');
        sheetView.addChild(selection);
        sheetViews.addChild(sheetView);
        worksheet.addChild(sheetViews);

        // sheetFormatPr
        var sheetFormatPr = Xml.createElement('sheetFormatPr');
        sheetFormatPr.set('baseColWidth', '8');
        sheetFormatPr.set('defaultRowHeight', '15');
        worksheet.addChild(sheetFormatPr);

        // sheetData
        var sheetData = Xml.createElement('sheetData');
        for (rowIndex in sheet.rows.keys()) {
            var row = sheet.rows[rowIndex];
            sheetData.addChild(xmlFromRow(row, rowIndex, sharedStrings));
        }
        worksheet.addChild(sheetData);

        return worksheet;
    }

    static private function xmlFromRow(row:Map<Int, Cell>, rowIndex:Int, sharedStrings:SharedStrings):Xml
    {
        var rowXml = Xml.createElement('row');
        rowXml.set('r', Std.string(rowIndex + 1));
        for (cellIndex in row.keys()) {
            var cell = row[cellIndex];
            if (cell.cellType == None) continue;
            var cellXml = xmlFromCell(cell, cellIndex, rowIndex, sharedStrings);
            rowXml.addChild(cellXml);
        }
        return rowXml;
    }

    static private function xmlFromCell(cell:Cell, cellIndex:Int, rowIndex:Int, sharedStrings:SharedStrings):Xml
    {
        var cellXml = Xml.createElement('c');
        cellXml.set('r', nameFromCoordinates(cellIndex, rowIndex));
        switch (cell.cellType) {
            case Text:
                cellXml.set('t', 's');
                var v = valueXmlFromSharedString(cell, sharedStrings);
                cellXml.addChild(v);
            case Int | Float:
                cellXml.set('t', 'n');
                var v = valueXmlFromNumber(cell);
                cellXml.addChild(v);
            case Date | Time:
                cellXml.set('t', 'd');
                cellXml.set('s', styleFromType(cell.cellType));
                var v = valueXmlFromDate(cell);
                cellXml.addChild(v);
            case Formula:
                addValueXmlFromFormula(cellXml, cell);
            case None:
                trace('err');
        }
        return cellXml;
    }

    static private function styleFromType(cellType:Cell.CellType):String
    {
        return switch (cellType) {
            case Date: '1';
            case Time: '2';
            default: '0';
        }
    }

    static private function valueXmlFromSharedString(cell:Cell, sharedStrings:SharedStrings):Xml
    {
        var v = Xml.createElement('v');
        var likeElements = sharedStrings[cell.value];
        if (likeElements == null) {
            sharedStrings[cell.value] = new Array<Xml>();
            likeElements = sharedStrings[cell.value];
        }
        likeElements.push(v);
        return v;
    }

    static private function valueXmlFromNumber(cell:Cell):Xml
    {
        var v = Xml.createElement('v');
        var number = Xml.createPCData(Std.string(cell.numberValue));
        v.addChild(number);
        return v;
    }

    static private function valueXmlFromDate(cell:Cell):Xml
    {
        var v = Xml.createElement('v');
        var timeString = DateTools.format(cell.dateValue, '%Y-%m-%dT%H:%M:%S.000');
        var number = Xml.createPCData(timeString);
        v.addChild(number);
        return v;
    }

    static private function addValueXmlFromFormula(cellXml:Xml, cell:Cell)
    {
        var v = Xml.createElement('v');
        var calculation = Xml.createPCData(Std.string(cell.numberValue));
        v.addChild(calculation);
        var f = Xml.createElement('f');
        var formula = Xml.createPCData(cell.value);
        f.addChild(formula);
        cellXml.addChild(f);
        cellXml.addChild(v);
    }

    static public function nameFromCoordinates(columnIndex:Int, rowIndex:Int):String
    {
        return columnLetterFromIndex(columnIndex) + Std.string(rowIndex + 1);
    }

    static private function columnLetterFromIndex(index:Int):String
    {
        // TODO validate between 0 and 18277
        var letters = new Array<String>();
        index++;
        while (index > 0) {
            var remainder = index % 26;
            index = Std.int(index / 26);
            if (remainder == 0) {
                remainder = 26;
                index -= 1;
            }
            letters.push(String.fromCharCode(64+remainder));
        }
        letters.reverse();
        return letters.join('');
    }

}