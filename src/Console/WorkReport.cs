using System;
using System.Collections.Generic;
using System.Drawing;

namespace kuujinbo.EPPlusWrapper
{
    public class WorkReport
    {
        const int SHEETS = 5;

        Queue<int?> _hours = new Queue<int?>();
        Queue<int?> _people = new Queue<int?>();
        Random _random = new Random();

        public WorkReport()
        {
            FillQueue(_hours, 1000, 1, 4);
            FillQueue(_people, 10000, 0, 8);
        }

        public byte[] CreateWorkReport()
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

                return writer.GetAllBytes();
            }
        }

        private void FillQueue(Queue<int?> queue, int times, int min, int max)
        {
            var val = 0;
            var tmp = 0;
            for (int i = 0; i < times; ++i)
            {
                while (tmp == val)
                {
                    tmp = _random.Next(min, max);
                }
                val = tmp;
                queue.Enqueue(val);
            }
        }

        #region sheet indexes / headings
        public const int ColumnAvail = 2;
        public const int ColumnReason = 3;
        public const int ColumnShiftLength = 5;
        public const int ColumnShiftName = 6;

        public static readonly string[] ShiftNames = { "DAY", "SWING", "GRAVE" };

        public static readonly string[] WorkHeadings =
        {
            "PROJECT/DEPT", "AVAILABILITY / WORK ITEM", "REASON", "SHOP / TRADE / CODE", "SHIFT LENGTH (Hours)", "SHIFT"
        };
        #endregion

        public int NumberOfDays { get; private set; }
        public IList<string> DayNames { get; private set; }
        public IList<Cell> Days { get; private set; }

        public void InitDays(DateTime start, DateTime end)
        {
            if (DayNames != null && Days != null) return;

            NumberOfDays = (end - start).Days;
            DayNames = new List<string>();
            Days = new List<Cell>();

            while (start <= end)
            {
                DayNames.Add(UpperTwoLetterDay(start));
                var daysCell = new Cell()
                {
                    Value = start.Day.ToString(),
                    AllBorders = true,
                    Bold = true,
                    HorizontalAlignment = CellAlignment.HorizontalCenter
                };
                if (start.Day == 1 || start.Day == 25)
                {
                    daysCell.BackgroundColor = Color.Red;
                }
                else if (start.DayOfWeek == DayOfWeek.Saturday || start.DayOfWeek == DayOfWeek.Sunday)
                {
                    daysCell.BackgroundColor = Color.LightBlue;
                }
                Days.Add(daysCell);

                start = start.AddDays(1);
            }
        }

        private string UpperTwoLetterDay(DateTime day)
        {
            return day.ToString("ddd").Substring(0, 2).ToUpper();
        }

        public void WriteProjectHeadingRow(ExcelWriter writer, int row, int col)
        {
            var cell = new Cell()
            {
                AllBorders = true,
                Bold = true,
                HorizontalAlignment = CellAlignment.HorizontalCenter
            };

            var headings = WorkReport.WorkHeadings;
            for (int i = 0; i < headings.Length; ++i)
            {
                cell.Value = headings[i];
                writer.WriteCell(row, i + 1, cell);
            }

            var hoursStart = headings.Length + 1;
            for (int i = 0; i <= NumberOfDays; ++i)
            {
                cell.Value = DayNames[i];
                writer.WriteCell(row, hoursStart++, cell);
            }

            cell.Value = "Total Hours";
            writer.WriteCell(row, hoursStart, cell);
        }


        public void WriteRequestData(ExcelWriter writer, int startRow, int hoursStartColumn, List<int> data)
        {
            var lastColumn = hoursStartColumn + NumberOfDays + 1;
            writer.FreezePanes(1, lastColumn);

            foreach (var d in data)
            {
                // empty merged cells
                writer.WriteMergedCell(
                    new CellRange(startRow, 1, hoursStartColumn - 1),
                    new Cell() { BackgroundColor = Color.LightGray, AllBorders = true }
                );

                // days of month
                for (int i = 0; i <= NumberOfDays; ++i) writer.WriteCell(startRow, hoursStartColumn + i, Days[i]);

                // empty summary column stores excel SUM formulas in later rows
                writer.WriteCell(
                    startRow,
                    lastColumn,
                    new Cell() { AllBorders = true, BackgroundColor = Color.LightGray }
                );

                ++startRow;

                // merged data cells
                for (int i = 1; i < ColumnShiftName; ++i)
                {
                    var cell = i != ColumnShiftName - 1
                        ? new Cell() { AllBorders = true, Value = "test" }
                        : new Cell() 
                        { 
                            AllBorders = true, Value = _hours.Dequeue(), 
                            NumberFormat = Cell.FORMAT_TWO_DECIMAL 
                        };
                    writer.WriteMergedCell(
                        new CellRange(startRow, i, startRow + 2, i),
                        cell
                    );
                }

                for (int i = 0; i < ShiftNames.Length; ++i)
                {
                    var currentRow = startRow + i;
                    writer.WriteCell(
                        currentRow, 
                        ColumnShiftName,
                        new Cell() { AllBorders = true, Value = ShiftNames[i] }
                    );

                    var days = new int[NumberOfDays];
                    var hoursCell = new Cell() 
                    { 
                        AllBorders = true, NumberFormat = Cell.FORMAT_WHOLE_NUMBER 
                    };
                    for (int j = 0; j <= NumberOfDays; ++j) 
                    {
                        var value = _people.Dequeue();
                        hoursCell.Value = value > 0 ? value : null;
                        writer.WriteCell(currentRow, ColumnShiftName + 1 + j, hoursCell);
                    }

                    // hours subtotal
                    var totalCell = new Cell()
                    {
                        AllBorders = true,
                        Formula = string.Format(
                            "{0}*{1}",
                            writer.GetRowSum(hoursStartColumn, lastColumn - 1, currentRow),
                            writer.GetAddress(currentRow, hoursStartColumn - 2)
                        ),
                        NumberFormat = Cell.FORMAT_TWO_DECIMAL
                    };
                    writer.WriteCell(currentRow, lastColumn, totalCell);
                }
                startRow += 3;
            }

            /* ---------------------------------------------------------------
             *  worksheet subtotals
             *  --------------------------------------------------------------
             */
            // people subtotals
            writer.WriteMergedCell(
                new CellRange(startRow, 1, ColumnShiftName),
                new Cell()
                {
                    AllBorders = true,
                    BackgroundColor = Color.LightGray,
                    Bold = true,
                    HorizontalAlignment = CellAlignment.HorizontalRight,
                    Value = "Total People"
                }
            );

            var hourSumCell = new Cell()
            {
                AllBorders = true,
                BackgroundColor = Color.LightGray,
                Bold = true,
                NumberFormat = Cell.FORMAT_WHOLE_NUMBER
            };
            for (int i = 0; i <= NumberOfDays; ++i)
            {
                var index = ColumnShiftName + 1 + i;
                hourSumCell.Formula = writer.GetColumnSum(1, startRow - 1, index);
                writer.WriteCell(startRow, index, hourSumCell);
            }
            writer.WriteCell(
                startRow,
                lastColumn,
                new Cell() { BackgroundColor = Color.LightGray, AllBorders = true }
            );

            ++startRow;

            // work hour sum
            writer.WriteMergedCell(
                new CellRange(startRow, 1, lastColumn - 1),
                new Cell()
                {
                    AllBorders = true,
                    BackgroundColor = Color.LightGray,
                    Bold = true,
                    HorizontalAlignment = CellAlignment.HorizontalRight,
                    Value = "Total Work Hours"
                }
            );

            writer.WriteCell(
                startRow,
                lastColumn,
                new Cell()
                {
                    AllBorders = true,
                    BackgroundColor = Color.LightGray,
                    Bold = true,
                    Formula = writer.GetColumnSum(1, startRow - 1, lastColumn),
                    NumberFormat = Cell.FORMAT_TWO_DECIMAL
                }
            );
        }
    }
}