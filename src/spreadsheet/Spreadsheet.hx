package spreadsheet;

import haxe.Json;

/**
 * A spreadsheet. A spreadsheet can contain a number of sheets.
 * This allows adding/removing sheets to a spreadsheet.
 */
class Spreadsheet
{
    private var sheets:Array<Sheet>;

    public function new()
    {
        sheets = new Array<Sheet>();
    }

    /**
     * Returns a new Spreadsheet created from given json
     */
    public static function fromJson(jsonString:String):Spreadsheet
    {
        var json = Json.parse(jsonString);
        var spreadsheet = new Spreadsheet();

        var dateExp:EReg = ~/^\d\d\d\d-\d\d-\d\d( \d\d:\d\d:\d\d)?$/;
        var timeExp:EReg = ~/^\d\d:\d\d:\d\d$/;

        for (n in Reflect.fields(json)) {
            var sheet = spreadsheet.sheetNamed(n);
            var rows:Array<Dynamic> = Reflect.field(json, n);
            var rowNum = 0;
            for (row in rows) {
                var cells:Array<Dynamic> = row;
                var cellNum = 0;
                for (cellMetadata in cells) {
                    var cell = sheet.cellAt(cellNum, rowNum);
                    switch (Type.typeof(cellMetadata)) {
                        case TNull: continue;
                        case TInt: cell.setInt(cellMetadata);
                        case TFloat: cell.setFloat(cellMetadata);
                        case TClass(String):
                            if (dateExp.match(cellMetadata)) {
                                cell.setDate(Date.fromString(cellMetadata));
                            }
                            else if (timeExp.match(cellMetadata)) {
                                // Should this hack of adding a date to get around weirdness in haxe live somewhere else?
                                var date = Date.fromString('1971-01-01 ' + cellMetadata);
                                cell.setTime(date);
                            }
                            else cell.setText(cellMetadata);
                        default: trace('Not implemented type found in json');
                    }
                    cellNum++;
                }
                rowNum++;
            }
        }

        return spreadsheet;
    }

    /**
     * Returns sheet at index without doing bounds checking
     * @param index 
     * @return Sheet
     */
    public function sheetAt(index:Int):Sheet
    {
        return sheets[index];
    }

    /**
     * Find a sheet named `name` and if it doesn't exist, add a new one at the end.
     * @return Sheet
     */
    public function sheetNamed(name:String):Sheet
    {
        for (sheet in sheets) {
            if (sheet.name == name) {
                return sheet;
            }
        }
        var sheet = new Sheet(name);
        sheets.push(sheet);
        return sheet;
    }

    /**
     * Returns number of sheets
     * @return Int
     */
    public function numSheets():Int
    {
        return sheets.length;
    }
}
