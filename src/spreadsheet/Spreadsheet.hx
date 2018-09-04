package spreadsheet;

typedef SheetEntry = {
    var sheet:Sheet;
    var name:String;
}

/**
 * A spreadsheet. A spreadsheet can contain a number of sheets.
 * This allows adding/removing sheets to a spreadsheet.
 */
class Spreadsheet
{
    private var sheets:Array<SheetEntry>;

    public function new()
    {
        sheets = new Array<SheetEntry>();
    }

    /**
     * Returns sheet at index without doing bounds checking
     * @param index 
     * @return Sheet
     */
    public function sheetAt(index:Int):Sheet
    {
        return sheets[index].sheet;
    }

    /**
     * Find a sheet named `name` and if it doesn't exist, add a new one at the end.
     * @return Sheet
     */
    public function sheetNamed(name:String):Sheet
    {
        for (sheetEntry in sheets) {
            if (sheetEntry.name == name) {
                return sheetEntry.sheet;
            }
        }
        var sheet = new Sheet();
        var sheetEntry = { sheet: sheet, name: name };
        sheets.push(sheetEntry);
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
