﻿/* ===========================================================================
 * __SIMPLE__ wrappers for EPPlus
 * ===========================================================================
 */
using System;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace kuujinbo.EPPlusWrapper
{
    #region enum
    /// <summary>print orientation</summary>
    public enum PageOrientation
    {
        Landscape = eOrientation.Landscape, 
        Portrait = eOrientation.Portrait
    }

    /// <summary>
    /// ISO 216 paper size. TODO: add members as needed
    /// </summary>
    public enum PrintSize
    {
        A4 = ePaperSize.A4
    }
    #endregion

    public class ExcelWriter : IDisposable
    {
        private ExcelPackage _package;
        public ExcelWriter()
        {
            _package = new ExcelPackage();
        }

        /// <summary>
        /// landscape by default
        /// </summary>
        public PageOrientation Orientation { get; set; }

        /// <summary>
        /// A4 by default
        /// </summary>
        public PrintSize PrintSize { get; set; }

        #region worksheet
        /// <summary>
        /// reference to current worksheet
        /// </summary>
        public ExcelWorksheet Worksheet { get; set; }

        /// <summary>
        /// Add named worksheet with layout; default orientation and print size
        /// </summary>
        public ExcelWriter AddSheet(string sheetName, bool pageLayoutView = false)
        {
            AddSheet(sheetName, PageOrientation.Landscape, PrintSize.A4, pageLayoutView);
            return this;
        }

        /// <summary>
        /// Add named worksheet with orientation, print size, and layout
        /// </summary>
        public ExcelWriter AddSheet(
            string sheetName, 
            PageOrientation orientation,
            PrintSize printSize,
            bool pageLayoutView = false)
        {
            Worksheet = _package.Workbook.Worksheets.Add(sheetName);
            Worksheet.PrinterSettings.Orientation = (eOrientation) orientation;
            Worksheet.PrinterSettings.PaperSize = (ePaperSize) printSize;
            Worksheet.View.PageLayoutView = pageLayoutView;
            return this;
        }

        /// <summary>
        /// set mirrored page margins
        /// </summary>
        public void SetMargins(decimal all)
        {
            SetMargins(all, all, all, all);
        }

        /// <summary>
        /// set mirrored top and bottom page margins
        /// </summary>
        public void SetMargins(decimal leftRight, decimal topBottom)
        {
            SetMargins(leftRight, leftRight, topBottom, topBottom);
        }

        /// <summary>
        /// set individual page margins
        /// </summary>
        public void SetMargins(decimal left, decimal right, decimal top, decimal bottom)
        {
            Worksheet.PrinterSettings.LeftMargin = left;
            Worksheet.PrinterSettings.RightMargin = right;
            Worksheet.PrinterSettings.BottomMargin = bottom;
            Worksheet.PrinterSettings.TopMargin = top;
        }

        /// <summary>
        /// set default worksheet styles: font size, family, alignment, and wrapping.
        /// see Cell class for defaults; wrapping **ALWAYS** on.
        /// </summary>
        public ExcelWriter SetWorkSheetStyles(int fontSize, string fontFamily = "Arial")
        {
            using (var allCells = Worksheet.Cells)
            {
                var style = allCells.Style;
                style.WrapText = true;
                style.HorizontalAlignment = (ExcelHorizontalAlignment)Cell.DefaultHorizontalAlignment;
                style.VerticalAlignment = (ExcelVerticalAlignment)Cell.DefaultVerticalAlignment;
                style.Font.Size = fontSize;
                style.Font.Name = fontFamily;
            }
            return this;
        }
        #endregion

        #region header/footer
        /// <summary>
        /// header and footer font family and weight
        /// </summary>
        public const string HEADER_FOOTER_FONT = "&\"Arial,Regular Bold\"";
        public const string PAGE_X_OF_Y = "&{0}Page {1} of {2}";

        /// <summary>
        /// set header text for one or more of it's three sections. 
        /// see GetHeaderFooterText() to set font size
        /// </summary>
        public void SetHeaderText(string left, string center, string right)
        {
            if (!string.IsNullOrWhiteSpace(left)) Worksheet.HeaderFooter.OddHeader.LeftAlignedText = left;
            if (!string.IsNullOrWhiteSpace(center)) Worksheet.HeaderFooter.OddHeader.CenteredText = center;
            if (!string.IsNullOrWhiteSpace(right)) Worksheet.HeaderFooter.OddHeader.RightAlignedText = right;
        }

        /// <summary>
        /// set footer text for one or more of it's three sections. 
        /// see GetHeaderFooterText() to set font size
        /// </summary>
        public void SetFooterText(string left, string center, string right)
        {
            if (!string.IsNullOrWhiteSpace(left)) Worksheet.HeaderFooter.OddFooter.LeftAlignedText = left;
            if (!string.IsNullOrWhiteSpace(center)) Worksheet.HeaderFooter.OddFooter.CenteredText = center;
            if (!string.IsNullOrWhiteSpace(right)) Worksheet.HeaderFooter.OddFooter.RightAlignedText = right;
        }

        /// <summary>
        /// EPPlus wrapper to properly format setting header/footer text
        /// </summary>
        public string GetHeaderFooterText(int fontSize, string text)
        {
            return string.Format("&{0}{1}{2}", fontSize, HEADER_FOOTER_FONT, text);
            //                                           ^^^^^^^^^^^^^^^^^^
            // font family and weight hard-coded for simplicity
        }

        /// <summary>
        /// standard Page 'X' of 'Y' header/footer
        /// </summary>
        public string GetPageNumOfTotalText(int fontSize)
        {
            return string.Format(
                PAGE_X_OF_Y,
                fontSize,
                ExcelHeaderFooter.PageNumber,
                ExcelHeaderFooter.NumberOfPages
            );
        }
        #endregion

        /// <summary>
        /// set worksheet column width
        /// </summary>
        public void SetColumnWidth(int col, double width)
        {
            Worksheet.Column(col).Width = width;
        }

        /// <summary>
        /// set all worksheet column widths at once
        /// </summary>
        public void SetColumnWidths(params double[] widths)
        {
            for (int i = 0; i < widths.Length; ++i)
            {
                Worksheet.Column(i + 1).Width = widths[i];
            }
        }

        /// <summary>
        /// required when creating Excel formulas; EPPlus wrapper to get 
        /// proprietary **Excel** cell address, versus accessing cell by index 
        /// </summary>
        public string GetAddress(int row, int col)
        {
            return Worksheet.Cells[row, col].Address;
        }

        /// <summary>
        /// string required to set a cell's formula for a **COLUMN**:
        /// e.g. SUM(A1:A4)
        /// </summary>
        public string GetColumnSum(int rowStart, int rowEnd, int col)
        {
            return string.Format(
                "SUM({0}:{1})",
                GetAddress(rowStart, col),
                GetAddress(rowEnd, col)
            );
        }

        /// <summary>
        /// string required to set a cell's formula for a **ROW**:
        /// e.g. SUM(A1:D4)
        /// </summary>
        public string GetRowSum(int colStart, int colEnd, int row)
        {
            return string.Format(
                "SUM({0}:{1})",
                GetAddress(row, colStart),
                GetAddress(row, colEnd)
            );
        }

        /// <summary>
        /// write **SINGLE** cell to current worksheet
        /// </summary>
        public void WriteCell(int row, int col, Cell cell)
        {
            Write(Worksheet.Cells[row, col], cell);
        }

        /// <summary>
        /// write cell range to current worksheet 
        /// </summary>
        public void WriteRange(CellRange cRange, Cell cell)
        {
            var range = Worksheet.Cells[
                cRange.FromRow,
                cRange.FromCol,
                cRange.ToRow,
                cRange.ToCol
            ];
            range.Merge = cRange.Merge;

            Write(range, cell);
        }

        /// <summary>
        /// write one or more cells to current worksheet
        /// </summary>
        private void Write(ExcelRange range, Cell cell)
        {
            using (range)
            {
                var style = range.Style;
                range.Value = cell.Value;
                range.Style.Font.Bold = cell.Bold;

                if (cell.AllBorders)
                {
                    style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                if (cell.BackgroundColor != Cell.DefaultBackgroundColor)
                {
                    style.Fill.PatternType = ExcelFillStyle.Solid;
                    style.Fill.BackgroundColor.SetColor(cell.BackgroundColor);
                }

                if (cell.FontColor != Cell.DefaultFontColor)
                {
                    style.Font.Color.SetColor(cell.FontColor);
                }

                if (cell.FontSize > Cell.MIN_FONT_SIZE)
                {
                    style.Font.Size = cell.FontSize;
                }

                if (cell.HorizontalAlignment != Cell.DefaultHorizontalAlignment)
                {
                    style.HorizontalAlignment = (ExcelHorizontalAlignment)cell.HorizontalAlignment;
                }

                if (cell.VerticalAlignment != Cell.DefaultVerticalAlignment)
                {
                    style.VerticalAlignment = (ExcelVerticalAlignment)cell.VerticalAlignment;
                }

                if (!string.IsNullOrWhiteSpace(cell.Formula))
                {
                    range.Formula = cell.Formula;
                }

                style.Numberformat.Format = string.IsNullOrWhiteSpace(cell.NumberFormat)
                    ? Cell.FORMAT_TEXT : cell.NumberFormat;
            }
        }

        /// <summary>
        /// method should be called when ready to send Excel to the client.
        /// </summary>
        public byte[] GetAllBytes()
        {
            AddBlankSheet();
            return _package.GetAsByteArray();
        }

        public const string NO_SHEETS_MESSAGE = "NO DATA AVAILABLE";
        public const int NO_SHEETS_END_COL = 20;

        /// <summary>
        /// add blank sheet with a simple message when workbook has no sheets.
        /// </summary>
        private void AddBlankSheet()
        {
            if (_package.Workbook.Worksheets.Count == 0)
            {
                AddSheet(NO_SHEETS_MESSAGE);
                WriteRange(
                    new CellRange(1, 1, NO_SHEETS_END_COL, true),
                    new Cell()
                    {
                        AllBorders = true,
                        Bold = true,
                        BackgroundColor = Color.Yellow,
                        FontSize = 20,
                        HorizontalAlignment = CellAlignment.HorizontalCenter,
                        Value = NO_SHEETS_MESSAGE
                    }
                );
            }
        }

        #region Dispose
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                _package.Dispose();
            }
        }
        #endregion
    }
}