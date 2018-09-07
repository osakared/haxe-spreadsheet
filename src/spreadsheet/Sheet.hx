package spreadsheet;

typedef CellIndex = {
    var rowIndex:Int;
    var columnIndex:Int;
}

class Sheet
{
    public var name(default, null):String;

    public var columnWidths(default, null):Map<Int, Float>;
    public var rowHeights(default, null):Map<Int, Float>;

    public var rows(default, null):Map<Int, Map<Int, Cell>>;

    public function new(_name:String)
    {
        name = _name;

        columnWidths = new Map<Int, Float>();
        rowHeights = new Map<Int, Float>();

        rows = new Map<Int, Map<Int, Cell>>();
    }

    public function cellAt(columnIndex:Int, rowIndex:Int):Cell
    {
        // TODO validate bounds here, return null if invalid
        var row = rows.get(rowIndex);
        if (row == null) {
            row = new Map<Int, Cell>();
            rows.set(rowIndex, row);
        }
        var cell = row.get(columnIndex);
        if (cell == null) {
            cell = new Cell();
            row.set(columnIndex, cell);
        }
        return cell;
    }
}