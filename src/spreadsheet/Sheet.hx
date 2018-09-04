package spreadsheet;

typedef CellIndex = {
    var rowIndex:Int;
    var columnIndex:Int;
}

class Sheet
{
    private var columnWidths:Map<Int, Float>;
    private var rowHeights:Map<Int, Float>;

    private var columns:Map<Int, Map<Int, Cell>>;

    public function new()
    {
        columnWidths = new Map<Int, Float>();
        rowHeights = new Map<Int, Float>();

        columns = new Map<Int, Map<Int, Cell>>();
    }

    public function cellAt(columnIndex:Int, rowIndex:Int):Cell
    {
        // TODO validate bounds here, return null if invalid
        var column = columns.get(columnIndex);
        if (column == null) {
            column = new Map<Int, Cell>();
            columns.set(columnIndex, column);
        }
        var cell = column.get(rowIndex);
        if (cell == null) {
            cell = new Cell();
            column.set(rowIndex, cell);
        }
        return cell;
    }
}