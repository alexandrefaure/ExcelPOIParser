using NPOI.SS.UserModel;

namespace TestExcelParser.Model
{
    internal class Cell
    {
        public CellType cellType;
        public int column;
        public string content;
        public string displayContent;
        public int row;
        public int sheet;
        public Style style;
    }

    internal class Style
    {
        public HorizontalAlignment alignment;
        public string backgroundColor;
        public string font;
        public double fontSize;
        public string foregroundColor;
        public bool isBold;
        public bool isItalic;
        public bool isStrikeout;
        public bool isUnderline;
        public VerticalAlignment verticalAlignment;
    }

    internal enum CellType
    {
        String,
        Numeric,
        Boolean,
        Date,
        Formula,
        Blank
    }

    internal class Sheet
    {
        public File file;
        public string name;
        public string position;
    }

    internal class File
    {
        public string name;
    }
}