using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RdlEngine.Render
{
    internal class WorkSheetSetting
    {
        public string WorksheetName;
        public ExcelWorksheet WorkSheet;

        public float? LastPixelsY;
        public int CurrentRow { get; private set; }
        public int CurrentCol { get; private set; }

        public List<string> MergeCells;

        public WorkSheetSetting(ExcelWorksheet excelWorkSheet)
        {
            // WorksheetName = worksheetName;
            WorkSheet = excelWorkSheet;
            StartCol();
            StartRow();
            MergeCells = new List<string>();
        }
        public WorkSheetSetting(string excelWorkSheetName)
        {
            WorksheetName = excelWorkSheetName;
            StartCol();
            StartRow();
            MergeCells = new List<string>();
        }


        public int NextCol()
        {
            return CurrentCol++;
        }
        public int NextCol(int increment)
        {
            return CurrentCol + increment;
        }
        public int StartCol()
        {
            return CurrentCol = 1;
        }
        public int StartRow()
        {
            return CurrentRow = 1;
        }
        public int NextRow()
        {
            return CurrentRow++;
        }
        public int NextRow(int increment)
        {
            return CurrentRow + increment;
        }


    }
}
