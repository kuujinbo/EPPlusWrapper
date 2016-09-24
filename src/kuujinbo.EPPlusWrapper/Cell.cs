/* ===========================================================================
 * __SIMPLE__ backing store wrapper for EPPlus and working w/Excel Cell
 * ===========================================================================
 */
using System.Drawing;
using OfficeOpenXml.Style;

namespace kuujinbo.EPPlusWrapper
{
    #region enum
    /// <summary>
    /// does **NOT** include all ExcelHorizontalAlignment/ExcelVerticalAlignment members
    /// </summary>
    public enum CellAlignment
    {
        HorizontalDefault = ExcelHorizontalAlignment.General,
        HorizontalLeft = ExcelHorizontalAlignment.Left,
        HorizontalCenter = ExcelHorizontalAlignment.Center,
        HorizontalRight = ExcelHorizontalAlignment.Right,
        VerticalTop = ExcelVerticalAlignment.Top,
        VerticalCenter = ExcelVerticalAlignment.Center,
        VerticalBottom = ExcelVerticalAlignment.Bottom
    }
    #endregion

    public class Cell
    {
        /* --------------------------------------------------------------------
         * convenience formats to set NumberFormat; 
         * __MUST__ be **VERY** careful when specifying formats, especially for
         * cells that perform calculations. 
         * --------------------------------------------------------------------
         */
        /// <summary>see NumberFormat property</summary>
        public const string FORMAT_TEXT = "@";
        /// <summary>see NumberFormat property</summary>
        public const string FORMAT_WHOLE_NUMBER = "#,##0";
        /// <summary>see NumberFormat property</summary>
        public const string FORMAT_TWO_DECIMAL = "#,##0.00";
        /// <summary>see NumberFormat property</summary>
        public const string FORMAT_CURRENCY = "$#,##0.00";

        /// <summary>Default worksheet background color</summary>
        public static readonly Color DefaultBackgroundColor = Color.White;
        /// <summary>Default worksheet font color</summary>
        public static readonly Color DefaultFontColor = Color.Black;
        /// <summary>Default worksheet horizontal cell alignment</summary>
        public static readonly CellAlignment DefaultHorizontalAlignment = CellAlignment.HorizontalDefault;
        /// <summary>Default worksheet vertical cell alignment</summary>
        public static readonly CellAlignment DefaultVerticalAlignment = CellAlignment.VerticalCenter;

        /// <summary>
        /// object is most straigtforward implementation to ensure that the
        /// cell data type is correctly parsed in Excel when Calculate() is 
        /// called from EPPlus
        /// </summary>
        public object Value { get; set; }

        public CellAlignment HorizontalAlignment { get; set; }
        public CellAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// current implementation is ALL or NONE.
        /// </summary>
        public bool AllBorders { get; set; }
        public bool Bold { get; set; }

        public const int MIN_FONT_SIZE = 4;
        /// <summary>
        /// any value less than MIN_FONT_SIZE is ignored
        /// </summary>
        public int FontSize { get; set; }

        /// <summary>
        /// if not familiar with System.Drawing.Color struct, see GetHtmlColor()
        /// </summary>
        public Color BackgroundColor { get; set; }
        /// <summary>
        /// if not familiar with System.Drawing.Color struct, see GetHtmlColor()
        /// </summary>
        public Color FontColor { get; set; }

        /// <summary>
        /// ColorTranslator.FromHtml() MS documentation is broken. HTML color 
        /// name, i.e. 'blue' **AND** hex color codes ('#ffffff') are allowed.
        /// </summary>
        public static Color GetHtmlColor(string webColor)
        {
            return ColorTranslator.FromHtml(webColor);
        }
        
        /// <summary>
        /// see EPPlus docs for supported formulas calculations and helper 
        /// methods in ExcelWriter()
        /// </summary>
        public string Formula { get; set; }

        /// <summary>
        /// **MUST** be **VERY** careful when specifying formats, especially
        /// for cells that perform calculations. 
        /// </summary>
        public string NumberFormat { get; set; }
        
        public Cell()
        {
            BackgroundColor = DefaultBackgroundColor;
            FontColor = DefaultFontColor;
            HorizontalAlignment = DefaultHorizontalAlignment;
            VerticalAlignment = DefaultVerticalAlignment;
        }
    }
}