using System;

namespace kuujinbo.EPPlusWrapper
{
    public class CellRange
    {
        public int FromRow { get; set; }
        public int ToRow { get; set; }
        public int FromCol { get; set; }
        public int ToCol { get; set; }
        public bool Merge { get; set; }

        /// <summary>
        /// get a range of cells for a **SINGLE** row, which also includes
        /// a **SINGLE* cell if needed
        /// </summary>
        public CellRange(int row, int fromCol, int toCol, bool merge = false)
            : this(row, fromCol, row, toCol, merge) { }

        /// <summary>
        /// get a range of cells that span **MORE** than one row
        /// </summary>
        public CellRange(int fromRow, int fromCol, int toRow, int toCol, bool merge = false)
        {
            FromRow = fromRow;
            FromCol = fromCol;
            ToRow = toRow;
            ToCol = toCol;
            Merge = merge;
        }
    }
}