package spreadsheet;

enum CellType
{
    None;
    Int;
    Text;
    Float;
    Date;
    Time;
    Formula;
}

enum VAlign
{
    VDefault;
    Top;
    Middle;
    Bottom;
}

enum HAlign
{
    HDefault;
    Left;
    Center;
    Right;
    Justified;
    Filled;
}

class Cell
{
    public var cellType(default, null):CellType;
    public var vAlign:VAlign;
    public var hAlign:HAlign;

    public var value(default, null):String;
    public var numberValue(default, null):Float;
    public var dateValue(default, null):Date;

    public function new()
    {
        cellType = None;
        vAlign = VDefault;
        hAlign = HDefault;
    }

    public function setText(newValue:String)
    {
        value = newValue;
        cellType = Text;
    }

    public function setFloat(newValue:Float)
    {
        numberValue = newValue;
        cellType = Float;
    }

    public function setDate(newValue:Date)
    {
        dateValue = newValue;
        cellType = Date;
    }

    public function setTime(newValue:Date)
    {
        dateValue = newValue;
        cellType = Time;
    }
}