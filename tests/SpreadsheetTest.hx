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
        sheet.cellAt(0, 0).setText('cell val');
        sheet.cellAt(0, 1).setInt(5);
        sheet.cellAt(0, 2).setInt(6);
        sheet.cellAt(2, 0).setDate(new Date(1989, 4, 2, 0, 0, 0));
        sheet.cellAt(2, 1).setTime(new Date(0, 0, 0, 16, 20, 0));
        sheet.cellAt(1, 0).setFormula('SUM(A:A)', 11);
        var output = File.write('test.xlsx');
        XlsxWriter.write(spreadsheet, output);
        output.close();
        return assert(true);
    }

}
