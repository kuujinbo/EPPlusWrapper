using Xunit;
using System;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace kuujinbo.EPPlusWrapper.Tests
{
    public class ExcelWriterTests : IDisposable
    {
        ExcelWriter _writer;
        public ExcelWriterTests()
        {
            _writer = new ExcelWriter();
            _writer.AddSheet("test");
        }

        public void Dispose()
        {
            _writer.Dispose();
        }

        [Fact]
        public void SetMargins_OneParam_SetsCorrectMargins()
        {
            var all = 0.1M;

            _writer.SetMargins(all);
            var result = _writer.Worksheet.PrinterSettings;

            Assert.Equal(all, result.LeftMargin);
            Assert.Equal(all, result.RightMargin);
            Assert.Equal(all, result.TopMargin);
            Assert.Equal(all, result.BottomMargin);
        }

        [Fact]
        public void SetMargins_TopBottom_SetsCorrectMargins()
        {
            var leftRight = 40M;
            var topBottom = 20M;

            _writer.SetMargins(leftRight, topBottom);
            var result = _writer.Worksheet.PrinterSettings;

            Assert.Equal(leftRight, result.LeftMargin);
            Assert.Equal(leftRight, result.RightMargin);
            Assert.Equal(topBottom, result.TopMargin);
            Assert.Equal(topBottom, result.BottomMargin);
        }

        [Fact]
        public void SetMargins_AllParams_SetsCorrectMargins()
        {
            var left = 4M;
            var right = 8M;
            var top = 20M;
            var bottom = 25M;

            _writer.SetMargins(left, right, top, bottom);
            var result = _writer.Worksheet.PrinterSettings;

            Assert.Equal(left, result.LeftMargin);
            Assert.Equal(right, result.RightMargin);
            Assert.Equal(top, result.TopMargin);
            Assert.Equal(bottom, result.BottomMargin);
        }

        [Fact]
        public void SetWorkSheetStyles_WithSizeAndFamily_SetsDefaults()
        {
            var expectedSize = 100;
            var expectedName = "Courier";

            _writer.SetWorkSheetStyles(expectedSize, expectedName);
            var defaultStyles = _writer.Worksheet.Cells.Style;

            Assert.Equal(expectedSize, defaultStyles.Font.Size);
            Assert.Equal(expectedName, defaultStyles.Font.Name);
        }

        [Fact]
        public void SetHeaderText_WithText_SetsTextInCorrectArea()
        {
            _writer.SetHeaderText("left", "center", "right");
            var result = _writer.Worksheet.HeaderFooter.OddHeader;

            Assert.Equal("left", result.LeftAlignedText);
            Assert.Equal("center", result.CenteredText);
            Assert.Equal("right", result.RightAlignedText);
        }

        [Fact]
        public void SetHeaderText_NullOrWhitespaceText_IsNoOp()
        {
            _writer.SetHeaderText(null, "       ", " ");
            var result = _writer.Worksheet.HeaderFooter.OddHeader;

            Assert.Null(result.LeftAlignedText);
            Assert.Null(result.CenteredText);
            Assert.Null(result.RightAlignedText);
        }
        
        [Fact]
        public void SetFooterText_WithText_SetsTextInCorrectArea()
        {
            _writer.SetFooterText("left footer", "center footer", "right footer");
            var result = _writer.Worksheet.HeaderFooter.OddFooter;

            Assert.Equal("left footer", result.LeftAlignedText);
            Assert.Equal("center footer", result.CenteredText);
            Assert.Equal("right footer", result.RightAlignedText);
        }

        [Fact]
        public void SetFooterText_NullOrWhitespaceText_IsNoOp()
        {
            _writer.SetFooterText(null, "       ", " ");
            var result = _writer.Worksheet.HeaderFooter.OddFooter;

            Assert.Null(result.LeftAlignedText);
            Assert.Null(result.CenteredText);
            Assert.Null(result.RightAlignedText);
        }

        [Fact]
        public void GetHeaderFooterText_SizeAndText_ReturnsFormattedText()
        {
            Assert.Equal(
                "&20&\"Arial,Regular Bold\"some text", 
                _writer.GetHeaderFooterText(20, "some text")
            );
        }

        [Fact]
        public void GetPageNumOfTotalText_SizeAndText_ReturnsFormattedText()
        {
            var fontSize = 20;

            _writer.Dispose();

            Assert.Equal(
                string.Format(
                    ExcelWriter.PAGE_X_OF_Y, 
                    fontSize, 
                    ExcelHeaderFooter.PageNumber,
                    ExcelHeaderFooter.NumberOfPages
                ),
                _writer.GetPageNumOfTotalText(fontSize)
            );
        }

        [Fact]
        public void SetColumnWidth_ColumnAndSize_SetsCorrectWidth()
        {
            var col = 10;
            var width = 4D;

            _writer.SetColumnWidth(col, width);
            var setColumn = _writer.Worksheet.Column(col);

            Assert.Equal(setColumn.Width, width);
        }

        [Fact]
        public void SetColumnWidths_WithDoubleParams_SetsCorrectWidths()
        {
            var widths = new double[] {4D, 8D, 20D};

            _writer.SetColumnWidths(widths);

            for (int i = 0; i < widths.Length; ++i)
            {
                // excel has one-based indexing 
                var column = _writer.Worksheet.Column(i + 1);
                Assert.Equal(column.Width, widths[i]);
            }
        }        

        [Fact]
        public void GetAddress_WithIndexParams_GetsExcelAddress()
        {
            Assert.Equal("A1", _writer.GetAddress(1, 1));
        }

        [Fact]
        public void GetColumnSum_WithIndexAndColumn_GetsFormattedForula()
        {
            Assert.Equal("SUM(A1:A4)", _writer.GetColumnSum(1, 4, 1));
        }

        [Fact]
        public void GetRowSum_WithIndexAndColumn_GetsFormattedForula()
        {
            Assert.Equal("SUM(B4:H4)", _writer.GetRowSum(2, 8, 4));
        }

        [Fact]
        public void GetAllBytes_WithSheet_GetsPackageBytes()
        {
            _writer = new ExcelWriter();
            var bytes = _writer.GetAllBytes();
            _writer.Dispose();

            Assert.IsType(typeof(byte[]), bytes);

            // verify AddBlankSheet() is called
            using (var ms = new MemoryStream(bytes))
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets[1];
                    var wrapperCell = sheet.Cells[1, 1];

                    Assert.Equal(1, package.Workbook.Worksheets.Count);
                    Assert.Equal(ExcelWriter.NO_SHEETS_MESSAGE, sheet.Name);
                    Assert.Equal(1, sheet.Dimension.End.Row);
                    Assert.Equal(ExcelWriter.NO_SHEETS_END_COL, sheet.Dimension.End.Column);
                    Assert.Equal(ExcelWriter.NO_SHEETS_MESSAGE, wrapperCell.Value);
                }
            }
        }

        [Fact]
        public void WriteCell_WithAddressAndCell_WritesToWorkSheet()
        {
            var cell1Value = 100d;
            var cell1 = new Cell()
            {
                AllBorders = true,
                Bold = true,
                BackgroundColor = Color.Black,
                FontColor = Color.Yellow,
                FontSize = 20,
                HorizontalAlignment = CellAlignment.HorizontalCenter,
                VerticalAlignment = CellAlignment.VerticalBottom,
                NumberFormat = Cell.FORMAT_TWO_DECIMAL,
                Value = cell1Value
            };
            var cell2 = new Cell() { Formula = "SUM(A1:B1)" };

            _writer.WriteCell(1, 1, cell1);
            _writer.WriteCell(1, 2, cell2);
            var bytes = _writer.GetAllBytes();
            _writer.Dispose();

            using (var ms = new MemoryStream(bytes))
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets[1];
                    var wrapperCell = sheet.Cells[1, 1];
                    var style = wrapperCell.Style;

                    Assert.Equal(1, package.Workbook.Worksheets.Count);
                    Assert.Equal(1, sheet.Dimension.End.Row);
                    Assert.Equal(2, sheet.Dimension.End.Column);
                    Assert.Equal(ExcelBorderStyle.Thin, style.Border.Left.Style);
                    Assert.Equal(ExcelBorderStyle.Thin, style.Border.Right.Style);
                    Assert.Equal(ExcelBorderStyle.Thin, style.Border.Top.Style);
                    Assert.Equal(ExcelBorderStyle.Thin, style.Border.Bottom.Style);
                    Assert.Equal(ExcelBorderStyle.Thin, style.Border.Bottom.Style);
                    Assert.Equal(ExcelFillStyle.Solid, style.Fill.PatternType);
                    Assert.True(style.Font.Bold);
                    Assert.Equal(cell1.FontSize, style.Font.Size);
                    Assert.Equal(
                        (ExcelHorizontalAlignment)cell1.HorizontalAlignment,
                        style.HorizontalAlignment
                    );
                    Assert.Equal(
                        (ExcelVerticalAlignment)cell1.VerticalAlignment,
                        style.VerticalAlignment
                    );
                    Assert.Equal(Cell.FORMAT_TWO_DECIMAL, wrapperCell.Style.Numberformat.Format);
                    Assert.Equal(cell1Value, wrapperCell.Value);

                    Assert.Equal(cell2.Formula, sheet.Cells[1, 2].Formula);
                }            
            }
        }
    }
}