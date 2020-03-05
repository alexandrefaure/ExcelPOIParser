using NPOI.SS.UserModel;

namespace TestExcelParser.Model
{
    public class Style
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
}