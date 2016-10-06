using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;

namespace kuujinbo.EPPlusWrapper
{
    class Program
    {
        static readonly string BASE_DIRECTORY = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        static void Main(string[] args)
        {
            var stopwatch = Stopwatch.StartNew();

            CreateReport();
            File.WriteAllBytes(
                Path.Combine(BASE_DIRECTORY, "epplus-test-work.xlsx"),
                new WorkReport().CreateWorkReport()
            );

            stopwatch.Stop();
            var timed = stopwatch.Elapsed;
            Console.WriteLine("run time: {0}.{1} seconds", timed.Seconds, timed.Milliseconds);
        }

        static void CreateReport()
        {
            // [1] create writer
            using (var writer = new ExcelWriter())
            {
                // [2] add worksheet to workbook w/optional parameters. put code from 
                //     here to step [5], **BEFORE** the writer.GetAllBytes() call in a
                //     repeating block to write more than one sheet
                writer.AddSheet(
                    "Sheet name", defaultColWidth: 10D, pageLayoutView: true
                );

                // [3] setup worksheet (ALL CALLS OPTIONAL, AND IN ANY ORDER)
                // set default font size
                writer.SetWorkSheetStyles(9);
                // set print margins
                writer.SetMargins(0.25M, 0.75M);
                // print table heading row(s) on every page
                writer.PrintRepeatRows(1, 3);
                // set left/center/right print header
                writer.SetHeaderText(
                    writer.GetHeaderFooterText(10, "Left"),
                    writer.GetHeaderFooterText(20, "Center"),
                    "Right"
                );
                // set left/center/right print footer
                writer.SetFooterText(null, writer.GetPageNumOfTotalText(8), null);

                // [4] write to current worksheet: 1-based index row and column
                // coordinates in **ANY** order.
                var cell = new Cell() 
                { 
                    AllBorders = true, Bold = true, BackgroundColor = Color.LightBlue
                };
                var headings = new string[] { "Heading 1", "Heading 2", "Sum"};
                var currentRow = 1;
                var colCount = headings.Length;

                // table 'heading'
                for (int i = 0; i < colCount; ++i)
                {
                    var colIndex = i + 1;
                    cell.Value = headings[i];
                    writer.WriteCell(currentRow, colIndex, cell);
                }

                // write data
                for (int i = 0; i < 100; ++i)
                {
                    ++currentRow;
                    cell = new Cell() { AllBorders = true };
                    var isInt = i % 2 == 0;
                    cell.NumberFormat = isInt ? Cell.FORMAT_WHOLE_NUMBER : Cell.FORMAT_CURRENCY;
                    for (int j = 1; j < colCount; ++j)
                    {
                        var val = i + j;
                        cell.Value = val;
                        writer.WriteCell(currentRow, j, cell);
                    }
                    cell.Formula = writer.GetRowSum(1, colCount - 1, currentRow);
                    writer.WriteCell(currentRow, colCount, cell);
                }

                // write arbitrary row
                cell = new Cell() 
                { 
                    AllBorders = true, BackgroundColor = Color.LightGray, Bold = true,
                    HorizontalAlignment = CellAlignment.HorizontalRight,
                    VerticalAlignment = CellAlignment.VerticalCenter,
                    Value = "Sum Last Column"
                };
                writer.WriteMergedCell(
                    new CellRange(++currentRow, 1, currentRow + 1, colCount - 1), 
                    cell
                );

                cell = new Cell()
                {
                    AllBorders = true, BackgroundColor = Color.LightGray, Bold = true,
                    Formula = writer.GetColumnSum(2, currentRow - 1, colCount),
                    VerticalAlignment = CellAlignment.VerticalCenter,
                    NumberFormat = Cell.FORMAT_TWO_DECIMAL
                };
                writer.WriteMergedCell(
                    new CellRange(currentRow, colCount, currentRow + 1, colCount),
                    cell
                );
        
                // [5] write workbook
                File.WriteAllBytes(
                    Path.Combine(BASE_DIRECTORY, "epplus-test-simple.xlsx"),
                    writer.GetAllBytes()
                );
            }
        }
    }
}