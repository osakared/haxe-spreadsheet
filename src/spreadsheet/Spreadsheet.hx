package spreadsheet;

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
