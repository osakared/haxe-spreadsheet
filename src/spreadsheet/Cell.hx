package spreadsheet;

enum CellType
{
    None;
    Text;
    Float;
    Date;
    Time;
    Formula;
}

class Cell
{
    public var cellType(default, null):CellType;
    public var value(default, null):String;

    public function new()
    {
        cellType = None;
    }

    public function setText(newValue:String)
    {
        value = newValue;
        cellType = Text;
    }
}