package spreadsheet.format;

import haxe.io.Output;
import haxe.Resource;
import haxe.zip.Entry;

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
            'xl/_rels/workbook.xml.rels',
            'xl/sharedStrings.xml',
            'xl/styles.xml',
            'xl/theme/theme1.xml',
            'xl/workbook.xml',
            'xl/worksheets/sheet1.xml'
        ];

        for (sheetIndex in 0...spreadsheet.numSheets()) {
            var sheet = spreadsheet.sheetAt(sheetIndex);
            var sheetXml = xmlFromSheet(sheet);
        }

        for (prefabFile in prefabFiles) {
            var bytes = Resource.getBytes('xlsx/' + prefabFile);
            var entry = {
                fileName: prefabFile,
                fileSize: bytes.length,
                fileTime: Date.now(),
                compressed: false,
                dataSize: bytes.length,
                data: bytes,
                crc32: haxe.crypto.Crc32.make(bytes),
                extraFields: null
            }
            entries.push(entry);
        }

        zipWriter.write(entries);
    }

    static private function xmlFromSheet(sheet:Sheet):Xml
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
        var rowIndex = 1;
        for (row in sheet.rows) {
            var rowXml = xmlFromRow(row);
            rowXml.set('r', Std.string(rowIndex));
            rowIndex++;
            sheetData.addChild(rowXml);
        }
        worksheet.addChild(sheetData);

        return worksheet;
    }

    static private function xmlFromRow(row:Map<Int, Cell>):Xml
    {
        var rowXml = Xml.createElement('row');
        for (cellIndex in row.keys()) {
            var cell = row[cellIndex];
            var cellXml = Xml.createElement('c');
        }
        return rowXml;
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