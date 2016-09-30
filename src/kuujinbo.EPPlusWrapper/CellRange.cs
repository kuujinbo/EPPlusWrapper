using System;

namespace kuujinbo.EPPlusWrapper
{
    public class CellRange
    {
        public int FromRow { get; set; }
        public int ToRow { get; set; }
        public int FromCol { get; set; }
        public int ToCol { get; set; }

        /// <summary>
        /// get cell range for **SINGLE** row, which includes **SINGLE* cell
        /// </summary>
        public CellRange(int row, int fromCol, int toCol)
            : this(row, fromCol, row, toCol) { }

        /// <summary>
        /// get cell range that span **MORE** than one row
        /// </summary>
        public CellRange(int fromRow, int fromCol, int toRow, int toCol)
        {
            FromRow = fromRow;
            FromCol = fromCol;
            ToRow = toRow;
            ToCol = toCol;
        }
    }
}