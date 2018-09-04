package;

import spreadsheet.Spreadsheet;
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

}
