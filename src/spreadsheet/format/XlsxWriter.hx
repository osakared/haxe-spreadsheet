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
}