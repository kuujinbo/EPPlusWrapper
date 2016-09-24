using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace kuujinbo.EPPlusWrapper
{
    class Program
    {
        const int SHEETS = 10;
        static readonly Random random = new Random();

        static void Main(string[] args)
        {
            var stopwatch = Stopwatch.StartNew();
            Init();
            stopwatch.Stop();
            var timed = stopwatch.Elapsed;
            Console.WriteLine("run time: {0}.{1} seconds", timed.Seconds, timed.Milliseconds);
        }

        static void Init()
        {
            var period = new CurtailmentReport();
            period.InitDays(new DateTime(2016, 12, 24), new DateTime(2017, 1, 2));

            using (var writer = new ExcelWriter())
            {
                for (int i = 0; i < SHEETS; ++i)
                {
                    var sheetName = string.Format("Project {0:D4}", i);
                    writer.AddSheet(sheetName, true).SetWorkSheetStyles(9);
                    writer.SetHeaderText(
                        null,
                        writer.GetHeaderFooterText(20, sheetName),
                        writer.GetPageNumOfTotalText(8)
                    );
                    writer.SetMargins(0.25M, 0.75M);

                    writer.SetColumnWidth(CurtailmentReport.ColumnAvail, 13);
                    writer.SetColumnWidth(CurtailmentReport.ColumnReason, 27);
                    writer.SetColumnWidth(CurtailmentReport.ColumnShiftLength, 8);
                    writer.SetColumnWidth(CurtailmentReport.ColumnShiftName, 8);

                    // set date header column widths
                    int hoursStartColumn = CurtailmentReport.ColumnShiftName + 1;
                    // int startColumn = COMMENT_COLUMN + 1;
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

                var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                File.WriteAllBytes(
                    Path.Combine(desktop, "epplus-test.xlsx"),
                    writer.GetAllBytes()
                );
            }
        }
    }
}