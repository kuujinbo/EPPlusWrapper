using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;

namespace kuujinbo.EPPlusWrapper
{
    class Program
    {
        const int SHEETS = 5;
        static readonly string BASE_DIRECTORY = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        static void Main(string[] args)
        {
            var stopwatch = Stopwatch.StartNew();

            CreateSimpleReport();
            CreateWorkReport();

            stopwatch.Stop();
            var timed = stopwatch.Elapsed;
            Console.WriteLine("run time: {0}.{1} seconds", timed.Seconds, timed.Milliseconds);
        }

        static void CreateSimpleReport()
        {
            // [1] create writer
            using (var writer = new ExcelWriter())
            {
                // [2] add worksheet to workbook w/optional parameters. put code from 
                //     here to step [5], **BEFORE** the writer.GetAllBytes() call in a
                //     repeating block to write more than one sheet
                writer.AddSheet(
                    "Sheet name", defaultColWidth: 4D, pageLayoutView: true
                );
                // [3] setup worksheet (ALL CALLS OPTIONAL, AND IN ANY ORDER)
                // set default font size
                writer.SetWorkSheetStyles(9);
                writer.SetHeaderText(
                    writer.GetHeaderFooterText(10, "Left"),
                    writer.GetHeaderFooterText(20, "Center"),
                    "Right"
                );
                writer.SetFooterText(
                    null,
                    writer.GetPageNumOfTotalText(8),
                    null
                );
                writer.SetMargins(0.25M, 0.75M);

                // [4] write to current worksheet: 1-based index row and column
                // coordinates in **ANY** order.
                var cell = new Cell() { AllBorders = true, Bold = true };
                cell.Value = "text    ";
                writer.WriteCell(1, 1, cell);

                cell.Value = 1000D;
                cell.NumberFormat = Cell.FORMAT_TWO_DECIMAL;
                writer.WriteCell(1, 2, cell);

                cell.Value = DateTime.Now;
                cell.NumberFormat = Cell.FORMAT_DATE;
                writer.WriteCell(1, 3, cell);

                cell.Value = "merged cell";
                cell.NumberFormat = Cell.FORMAT_TEXT;
                writer.WriteMergedCell(new CellRange(2, 1, 4, 8), cell);

                // [5] write workbook
                File.WriteAllBytes(
                    Path.Combine(BASE_DIRECTORY, "epplus-test-simple.xlsx"),
                    writer.GetAllBytes()
                );
            }
        }

        static void CreateWorkReport()
        {
            var period = new WorkReport();
            period.InitDays(new DateTime(2016, 12, 24), new DateTime(2017, 1, 2));

            using (var writer = new ExcelWriter())
            {
                for (int i = 0; i < SHEETS; ++i)
                {
                    var sheetName = string.Format("Project {0:D4}", i);
                    writer.AddSheet(sheetName, pageLayoutView: true);
                    writer.SetWorkSheetStyles(9);
                    writer.SetHeaderText(
                        null,
                        writer.GetHeaderFooterText(20, sheetName),
                        writer.GetPageNumOfTotalText(8)
                    );
                    writer.SetMargins(0.25M, 0.75M);

                    writer.SetColumnWidth(WorkReport.ColumnAvail, 13);
                    writer.SetColumnWidth(WorkReport.ColumnReason, 27);
                    writer.SetColumnWidth(WorkReport.ColumnShiftLength, 8);
                    writer.SetColumnWidth(WorkReport.ColumnShiftName, 8);

                    // set date header column widths
                    int hoursStartColumn = WorkReport.ColumnShiftName + 1;
                    int hoursEndColumn = hoursStartColumn + period.DayNames.Count;
                    for (int col = hoursStartColumn; col < hoursEndColumn; ++col)
                    {
                        writer.SetColumnWidth(col, 5);
                    }

                    period.WriteProjectHeadingRow(writer, 1, hoursStartColumn);

                    var testData = new List<int>();
                    for (int j = 1; j < 21; ++j) testData.Add(j);
                    period.WriteRequestData(writer, 2, hoursStartColumn, testData);
                }
                
                File.WriteAllBytes(
                    Path.Combine(BASE_DIRECTORY, "epplus-test-work.xlsx"),
                    writer.GetAllBytes()
                );
            }
        }
    }
}