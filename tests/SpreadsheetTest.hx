package;

import spreadsheet.Spreadsheet;
import spreadsheet.format.XlsWriter;
import spreadsheet.format.XlsxWriter;
import sys.io.File;
import tink.unit.Assert.*;

@:asserts
class SpreadsheetTest {

    public function new()
    {
    }

    public function testAdd()
    {
        var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.sheetNamed('Data');
        var cell = sheet.cellAt(0, 0);
        cell.setText('cell val');
        cell = sheet.cellAt(0, 0);
        return assert(cell.value == 'cell val');
    }

    public function testCoordinates()
    {
        return assert(XlsxWriter.nameFromCoordinates(4, 3) == 'E4');
    }

    public function testCreateXls()
    {
        var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.sheetNamed('Data');
        var cell = sheet.cellAt(0, 0);
        cell.setText('cell val');
        var output = File.write('test.xls');
        XlsWriter.write(spreadsheet, output);
        output.close();
        return assert(true);
    }

    public function testCreateXlsX()
    {
        var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.sheetNamed('Data');
        var cell = sheet.cellAt(0, 0);
        cell.setText('cell val');
        var output = File.write('test.xlsx');
        XlsxWriter.write(spreadsheet, output);
        output.close();
        return assert(true);
    }

}
