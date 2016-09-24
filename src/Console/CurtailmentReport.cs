﻿using System;
using System.Collections.Generic;
using System.Drawing;

namespace kuujinbo.EPPlusWrapper
{
    public class CurtailmentReport
    {
        #region common
        public const int COL_START = 1;
        #endregion

        #region project / code
        public const int ColumnAvail = 2;
        public const int ColumnReason = 3;
        public const int ColumnShiftLength = 5;
        public const int ColumnShiftName = 6;


        public const string DAY = "DAY";
        public const string SWING = "SWING";
        public const string GRAVE = "GRAVE";

        public static readonly string[] ShiftNames = { DAY, SWING, GRAVE };

        public static readonly string[] ProjectHeadings =
        {
            "PROJECT/DEPT", "AVAILABILITY / WORK ITEM", "REASON", "SHOP / TRADE / CODE", "SHIFT LENGTH (Hours)", "SHIFT"
        };

        /// <summary>
        /// 
        /// </summary>
        public static readonly int ProjectColumns = ProjectHeadings.Length;
        #endregion

        #region summary
        public static readonly string[] SummaryHeadings = 
        {
            "Project/Dept", "Total Work Hours", "Total Man Days", "Comment", "DH ConcurredBy"
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

            var headings = CurtailmentReport.ProjectHeadings;
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

        /// <summary>
        /// write & format numeric curtailment period days of month; range
        /// of cells written once for every OT request for project/department
        /// </summary>
        public void WriteRequestData(ExcelWriter writer, int startRow, int hoursStartColumn, List<int> data)
        {
            var lastColumn = hoursStartColumn + NumberOfDays + 1;
            foreach (var d in data)
            {
                // empty merged cells
                writer.WriteRange(
                    new CellRange(startRow, 1, hoursStartColumn - 1),
                    new Cell() { BackgroundColor = Color.LightGray, AllBorders = true }
                );

                // curtailment days of month
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
                        : new Cell() { AllBorders = true, Value = 8, NumberFormat = Cell.FORMAT_TWO_DECIMAL };
                    writer.WriteRange(
                        new CellRange(startRow, i, startRow + 2, i, true),
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

                    Random random = new Random();
                    var days = new int[NumberOfDays];
                    var hoursCell = new Cell() 
                    { 
                        AllBorders = true, NumberFormat = Cell.FORMAT_WHOLE_NUMBER 
                    };
                    for (int j = 0; j <= NumberOfDays; ++j) 
                    {
                        var value = random.Next(0, 2);
                        if (value > 0) hoursCell.Value = value;
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
            writer.WriteRange(
                new CellRange(startRow, 1, ColumnShiftName, true),
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
            writer.WriteRange(
                new CellRange(startRow, 1, lastColumn - 1, true),
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