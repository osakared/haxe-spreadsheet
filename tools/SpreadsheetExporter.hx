package;

import spreadsheet.Spreadsheet;
import spreadsheet.format.XlsxWriter;
import sys.io.File;

class SpreadsheetExporter
{
    static function main()
    {
        var args = Sys.args();
        if (args.length < 2) {
            trace('Usage: SpreadsheetExporter JSON_FILE SPREADSHEET_FILE');
            Sys.exit(1);
        }
        var jsonFilename = args[0];
        var spreadsheetFilename = args[1];

        var json = File.getContent(jsonFilename);
        var spreadsheet = Spreadsheet.fromJson(json);
        var output = File.write(spreadsheetFilename);
        XlsxWriter.write(spreadsheet, output);
        output.close();
    }
}