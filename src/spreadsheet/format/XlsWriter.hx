package spreadsheet.format;

import haxe.i18n.Utf16;
import haxe.io.Bytes;
import haxe.io.BytesOutput;
import haxe.io.Output;

class XlsWriter
{
    static public function write(spreadsheet:Spreadsheet, mainOutput:Output)
    {
        var output = new BytesOutput();
        output.bigEndian = false;
        // BOF 0x0809
        var bof = Bytes.ofData(cast [
            0x00, 0x06, 0x05, 0x00, 0xF2, 0x15, 0xCC, 0x07,
            0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00
        ]);
        writeRecord(0x0809, bof, output);

        // WINDOW1 0x003D
        var window1 = Bytes.ofData(cast [
            0x68, 0x01, 0x0E, 0x01, 0x5C, 0x3A, 0xBE, 0x23,
            0x38, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
            0x58, 0x02
        ]);
        writeRecord(0x003D, window1, output);
        // write required XF records
        writeXFs(output);
        // write a BOUNDSHEET record for each table
        var sheetRefOffsets = new Array<Int>();
        for (i in 0...spreadsheet.numSheets()) {
            var sheet = spreadsheet.sheetAt(i);
            // record offset to reference to sheet streams for future filling
            var sheetRefOffset = writeBoundsheet(sheet.name, output);
            sheetRefOffsets.push(sheetRefOffset);
        }
        // EOF
        writeRecordHeader(0x000A, 0, output);
        output.writeByte(0);
        // write each table as a sheet stream
        for (i in 0...spreadsheet.numSheets()) {
            // fill reference to this sheet first
            // TODO somehow I have to do this
            // write4bytes(out.size(), out.data() + sheetRefOffsets[i]);
            var sheet = spreadsheet.sheetAt(i);
            writeSheet(sheet, output);
        }

        output.close();
        mainOutput.write(output.getBytes());
    }

    static private function writeRecord(id:Int, data:Bytes, output:Output)
    {
        writeRecordHeader(id, data.length, output);
        output.write(data);
    }

    static private function writeRecordHeader(id:Int, length:Int, output:Output)
    {
        // output.writeByte(id & 0xFF);
        // output.writeByte((id >> 8) & 0xFF);
        output.writeUInt16(id);
        output.writeUInt16(length);
    }

    static private function writeBoundsheet(sheetName:String, output:BytesOutput):Int
    {
        // BOUNDSHEET 0x0085
        // Offset   Size    Contents
        // 0        4       Sheet's BOF stream position, 0 is OK
        // 4        1       Visibility: 0 = visible
        // 5        1       Sheet type: 0 = worksheet
        // 6        var     Sheet name: UTF16 string, 8-bit string length
        var utf16 = new Utf16(sheetName).toBytes();
        var length:Int = 4 + 1 + 1 + 2 + utf16.length * 2;
        writeRecordHeader(0x0085, length, output);
        var retval = output.length;
        output.writeInt32(0);
        output.writeByte(0);
        output.writeByte(0);
        output.writeByte(utf16.length);    // string length (8 bits)
        output.writeByte(1);               // option flags
        output.write(utf16);
        return retval;
    }

    static private function writeXFs(output:Output)
    {
        // XF 0x00E0
        // Offset   Size    Contents
        // 0        2       Index to FONT record
        // 2        2       Index to FORMAT record
        // 4        2       Bit     Mask    Contents
        //                  2-0     0007H   XF_TYPE_PROT � XF type, cell prot
        //                  15-4    FFF0H   Index to parent style XF
        //                                  (always FFFH in style XFs)
        // 6        1       Bit     Mask    Contents
        //                  2-0     07H     XF_HOR_ALIGN � Horizontal alignment
        //                  3       08H     1 = Text is wrapped at right border
        //                  6-4     70H     XF_VERT_ALIGN � Vertical alignment
        // 7        1       XF_ROTATION: Text rotation angle
        // 8        1       Indentation and direction
        // 9        1       Bit     Mask    Contents
        //                  7-2     FCH     XF_USED_ATTRIB � Used attributes
        // 10       4       Cell border lines and background area
        // 14       4       Various graphics options
        // 18       2       Colors
        var xf = Bytes.ofData(cast [
            0x00, 0x00, 0x00, 0x00, 0xF5, 0xFF, 0x20, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00
        ]);
        // write 16 basic records
        for (i in 0...16) {
            writeRecord(0x00E0, xf, output);
        }
        // write XF records that are going to be used
        var format_offset:Int = 2;
        var align_offset:Int = 6;
        var format_general:Int = 0;
        var format_date:Int = 14;
        var format_time:Int = 21;
        var halign_default:Int = 0x00;
        var halign_left:Int = 0x01;
        var halign_centered:Int = 0x02;
        var halign_right:Int = 0x03;
        var halign_justified:Int = 0x05;
        var halign_filled:Int = 0x04;
        var valign_top:Int = 0x00;
        var valign_middle:Int = 0x10;
        var valign_bottom:Int = 0x20;
        var formats:Array<Int> = [format_general, format_date, format_time];
        var hAligns:Array<Int> = [halign_default, halign_left, halign_centered,
            halign_right, halign_justified, halign_filled];
        var vAligns:Array<Int> = [valign_bottom, valign_top, valign_middle];
        
        xf.setUInt16(4, 0); // XF_TYPE_PROT = 0, index to parent style XF = 0

        for (format in formats) {
            for (hAlign in hAligns) {
                for (vAlign in vAligns) {
                    xf.setUInt16(format_offset, format);
                    xf.set(align_offset, hAlign | vAlign);
                    writeRecord(0x00E0, xf, output);
                }
            }
        }
    }

    static private function writeSheet(sheet:Sheet, output:Output)
    {
        // BOF 0x0809
        var bof = Bytes.ofData(cast [
            0x00, 0x06, 0x10, 0x00, 0xF2, 0x15, 0xCC, 0x07,
            0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00
        ]);
        writeRecord(0x0809, bof, output);
        // columns
        writeColumns(sheet, output);
        // rows
        writeRows(sheet, output);
        // EOF
        writeRecordHeader(0x000A, 0, output);
        output.writeByte(0);
    }

    static private function writeColumns(sheet:Sheet, output:Output)
    {
        for (column in sheet.columnWidths.keys()) {
            var columnWidth = sheet.columnWidths[column];
            // COLINFO 0x007D
            // Offset   Size    Contents
            // 0        2       Index to first column in the range
            // 2        2       Index to last column in the range
            // 4        2       Width of the columns in 1/256 of the width of
            //                  the zero character, using default font (first
            //                  FONT record in the file)
            // 6        2       Index to XF record for default column formatting
            // 8        2       Option flags:
            //                  Bits    Mask    Contents
            //                  0       0001H   1 = Columns are hidden
            //                  10-8    0700H   Outline level of the columns
            //                                  (0 = no outline)
            //                  12      1000H   1 = Columns are collapsed
            // 10       2       Not used
            var colinfo = Bytes.ofData(cast [
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0F, 0x00,
                0x06, 0x00, 0x00, 0x00
            ]);
            colinfo.setUInt16(0, column);
            colinfo.setUInt16(2, column);
            var w:Int = Std.int(columnWidthUnits(columnWidth) * 0x100);
            colinfo.setUInt16(4, w);
            writeRecord(0x007D, colinfo, output);
        }
    }

    static private function columnWidthUnits(width:Float):Float
    {
        return width * 42.85546875 / 225;
    }

    static private function writeRows(sheet:Sheet, output:Output)
    {
        for (rowKey in sheet.rowHeights.keys()) {
            var rowHeight = sheet.rowHeights[rowKey];

            var row = sheet.rows[rowKey];
            var firstCell:Null<Int> = null;
            var lastCell:Null<Int> = null;
            if (row != null) {
                for (cell in row.keys()) {
                    if (firstCell == null || cell < firstCell) firstCell = cell;
                    if (lastCell == null || cell > lastCell) lastCell = cell;
                }
            }
            if (firstCell == null) firstCell = 0;
            if (lastCell == null) lastCell = 0xffff;
            // ROW 0x0208
            // Offset   Size    Contents
            // 0        2       Index of this row
            // 2        2       Index to column of the first cell
            // 4        2       Index to column of the last cell + 1
            // 6        2       Bit     Mask    Contents
            //                  14-0    7FFFH   Height of the row, in twips
            //                                  (= 1/20 of a point)
            //                  15      8000H   0 = Row has custom height;
            //                                  1 = Row has default height
            // 8        2       Not used
            // 10       2       Not used
            // 12       4       Option flags and default row formatting:
            //                  0x00000100 works just fine. 0x04 should be
            //                  added to apply custom height
            var rowHeader = Bytes.ofData(cast [
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x80,
                0x06, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00
            ]);
            rowHeader.setUInt16(0, rowKey);
            rowHeader.setUInt16(2, firstCell);
            rowHeader.setUInt16(4, lastCell + 1);
            // Based on the current code, this will always be true, but maybe we'll change the code
            if (rowHeight != null) {
                rowHeader.setUInt16(6, Std.int(rowHeight * 20) & 0x7FFF);
                rowHeader.setInt32(12, 0x00000140);
            }
            writeRecord(0x0208, rowHeader, output);
        }

        // now the cells should go
        for (rowIndex in sheet.rows.keys()) {
            var row = sheet.rows[rowIndex];
            for (columnIndex in row.keys()) {
                var cell = row[columnIndex];
                if (cell.cellType == None) continue;
                writeCell(cell, columnIndex, rowIndex, output);
            }
        }
    }

    private static function writeCell(cell:Cell, column:Int, row:Int, output:Output)
    {
        if (cell.cellType == Text) {
            // LABEL 0x0204
            // Offset   Size    Contents
            // 0        2       Index to row
            // 2        2       Index to column
            // 4        2       Index to XF record
            // 6        var.    Unicode string, 16-bit string length
            var utf16 = new Utf16(cell.value).toBytes();
            writeRecordHeader(0x0204, 6 + 3 + utf16.length * 2, output);
            output.writeUInt16(row);
            output.writeUInt16(column);
            output.writeUInt16(xfIndex(cell));
            output.writeUInt16(utf16.length);   // string length (16 bits)
            output.writeByte(1);                // option flags
            output.write(utf16);
            return;
        }
        if (cell.cellType == Int || cell.cellType == Float) {
            // NUMBER 0x0203 (long)
            // Offset   Size    Contents
            // 0        2       Index to row
            // 2        2       Index to column
            // 4        2       Index to XF record
            // 6        8       IEEE 754 floating-point value
            //                  (64-bit double precision)
            writeRecordHeader(0x0203, 14, output);
            output.writeUInt16(row);
            output.writeUInt16(column);
            output.writeUInt16(xfIndex(cell));
            output.writeDouble(cell.numberValue);
            return;
        }
        if (cell.cellType == Date) {
            // NUMBER 0x0203 (date)
            // Offset   Size    Contents
            // 0        2       Index to row
            // 2        2       Index to column
            // 4        2       Index to XF record
            // 6        8       IEEE 754 floating-point value
            //                  (64-bit double precision)
            writeRecordHeader(0x0203, 14, output);
            output.writeUInt16(row);
            output.writeUInt16(column);
            output.writeUInt16(xfIndex(cell));
            output.writeDouble(excelDate(cell.dateValue));
            return;
        }
        if (cell.cellType == Time) {
            // NUMBER 0x0203 (time)
            // Offset   Size    Contents
            // 0        2       Index to row
            // 2        2       Index to column
            // 4        2       Index to XF record
            // 6        8       IEEE 754 floating-point value
            //                  (64-bit double precision)
            writeRecordHeader(0x0203, 14, output);
            output.writeUInt16(row);
            output.writeUInt16(column);
            output.writeUInt16(xfIndex(cell));
            output.writeDouble(excelTime(cell.dateValue));
            return;
        }
        // TODO implement Formula.. here's the original C++ code
        // if (cell.cellType == Formula) {
        //     // FORMULA 0x0006
        //     // Offset   Size    Contents
        //     // 0        2       Index to row
        //     // 2        2       Index to column
        //     // 4        2       Index to XF record
        //     // 6        8       Result of the formula
        //     // 14       2       Option flags: 0x0002 = Calculate on open
        //     // 16       4       Not used
        //     // 20       var     Formula data (RPN token array)
        //     int c1, r1, c2, r2;
        //     Formulas::parse(cell.getFormula(), c1, r1, c2, r2);
        //     byte FORMULA[] = {
        //         0x0D, 0x00, 0x25, 0x01, 0x00, 0x02, 0x00, 0x03,
        //         0xC0, 0x04, 0xC0, 0x19, 0x10, 0x00, 0x00};
        //     write2bytes((ushort)r1, FORMULA + 3);
        //     write2bytes((ushort)r2, FORMULA + 5);
        //     FORMULA[7] = (byte)c1;
        //     FORMULA[9] = (byte)c2;
        //     ushort len = (ushort)(6 + 8 + 2 + 4 + sizeof(FORMULA));
        //     writeRecordHeader(0x0006, len, out);
        //     write2bytes(row, out);
        //     write2bytes(col, out);
        //     write2bytes(xfIndex(cell), out);
        //     writeDouble(0, out);
        //     write2bytes(0x02, out);
        //     write4bytes(0, out);
        //     out.write(FORMULA, sizeof(FORMULA));
        //     return;
        // }
    }

    private static function xfIndex(cell:Cell):Int
    {
        var i:Int = 0;
        switch (cell.cellType) {
            case Date:        i = 1;
            case Time:        i = 2;
            default:          i = 0;
        }
        var j:Int = 0;
        switch (cell.hAlign) {
            case Left:        j = 1;
            case Center:      j = 2;
            case Right:       j = 3;
            case Justified:   j = 4;
            case Filled:      j = 5;
            default:          j = 0;
        }
        var k:Int = 0;
        switch (cell.vAlign) {
            case Top:         k = 1;
            case Middle:      k = 2;
            default:          k = 0;
        }
        return i * 3 * 6 + j * 3 + k;
    }

    private static function excelDate(date:Date):Int
    {
        var nYear = date.getFullYear();
        var nMonth = date.getMonth();
        var nDay = date.getDay();

        // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
        // leap year, but Excel/Lotus 123 think it is...
        if (nDay == 29 && nMonth == 2 && nYear == 1900) {
            return 60;
        }

        // DMY to Modified Julian calculatie with an extra substraction of 2415019.
        var nSerialDate:UInt = 
            Std.int((1461 * (nYear + 4800 + Std.int((nMonth - 14) / 12))) / 4) +
            Std.int((367 * (nMonth - 2 - 12 * ((nMonth - 14) / 12))) / 12) -
            Std.int((3 * (Std.int((nYear + 4900 + Std.int((nMonth - 14) / 12)) / 100))) / 4)
            + nDay - 2415019 - 32075;

        if (nSerialDate < 60) {
            // Because of the 29-02-1900 bug, any serial date 
            // under 60 is one off... Compensate.
            nSerialDate--;
        }

        return nSerialDate;
    }

    private static function excelTime(time:Date):Float
    {
        var h = time.getHours();
        var m = time.getMinutes();
        var s = time.getSeconds();
        // TODO do something about the lack of milis!
        var ms = 0; // time.getMillis();
        var msInDay:Float = 24 * 60 * 60 * 1000;
        return (ms + s * 1000 + m * 60 * 1000 + h * 60 * 60 * 1000) / msInDay;
    }
}
