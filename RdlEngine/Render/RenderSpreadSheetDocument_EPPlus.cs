
/*
 * 
 Copyright (C) 2004-2008  fyiReporting Software, LLC
 Copyright (C) 2011  Peter Gill <peter@majorsilence.com>
 Copyright (c) 2010 devFU Pty Ltd, Josh Wilson and Others (http://reportfu.org)
 Copyright (c) 2016 Daniel Romanowski http://dotlink.pl




 This file has been modified with suggestiong from forum users.
 *Obtained from Forum, User: sinnovasoft http://www.fyireporting.com/forum/viewtopic.php?t=1049

  This file is part of the fyiReporting RDL project.
	
   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

	   http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

using System;
using fyiReporting.RDL;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using fyiReporting.RDL.Utility;
using System.Security;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using RdlEngine.Render;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace fyiReporting.RDL
{


    ///<summary>
    /// Renders a report to SpreadSheet.
    ///</summary>
    [SecuritySafeCritical]
    internal class RenderSpreadSheetDocument_EPPlus : RenderBase
    {
        #region private

        ExcelPackage _excelPackage;

        MemoryStream _ms;
        List<WorkSheetSetting> _workSheetSettings;

        WorkSheetSetting _currentWorkSheetSetting;
        #endregion

        

        #region ctor
        public RenderSpreadSheetDocument_EPPlus(Report report, IStreamGen sg) : base(report, sg)
        {
            _ms = new MemoryStream();
            _workSheetSettings = new List<WorkSheetSetting>();

        }
        #endregion

        #region implementation abstract methods
        protected internal override void CreateDocument()
        {
            Report r = base.Report();

            _excelPackage = new ExcelPackage(_ms);

            // add a new worksheet to the empty workbook
            ExcelWorksheet worksheet =_excelPackage.Workbook.Worksheets.Add(r.Name);
            _workSheetSettings.Add(new WorkSheetSetting(worksheet));
            _currentWorkSheetSetting = _workSheetSettings.Last();
 
        }

      
        protected internal override void EndDocument(Stream sg)
        {
         
          /*        
           foreach (var worksheetSetting in _workSheetSettings)
            {
                Worksheet worksheet = _openXmlExportHelper.GetWorksheet(_spreadSheet, worksheetSetting.WorksheetName);
                foreach (var mergeCell in worksheetSetting.MergeCells)
                     _openXmlExportHelper.MergeCells(worksheet, mergeCell);
                
            }
            */
           
            _excelPackage.Save();

            byte[] contentbyte = _ms.ToArray();
            sg.Write(contentbyte, 0, contentbyte.Length);
            _ms.Dispose();
       
        }

       
        #endregion

        #region overrider virtual methods
        public override bool IsPagingNeeded()
        {
            return false;
        }
        public override void Textbox(Textbox tb, string t, Row row)
        {
            base.Textbox(tb, t, row);
            bool tableCell = IsTableCell(tb);
            if (!tableCell)
            {
                if (_currentWorkSheetSetting.LastPixelsY == null 
                 || _currentWorkSheetSetting.LastPixelsY != tb.Top.PixelsY)
                {
                        _currentWorkSheetSetting.LastPixelsY = tb.Top.PixelsY;
                        _currentWorkSheetSetting.NextRow();
                }
                
            }

            GetStyledValue(tb.Style, 
                            row,
                            _currentWorkSheetSetting.WorkSheet.Cells[_currentWorkSheetSetting.CurrentRow,_currentWorkSheetSetting.CurrentCol],
                            t);

            
            if (!tableCell)
            {
                _currentWorkSheetSetting.NextRow();
            }
           


        }

        public override bool TableStart(Table t, Row row)
        {
            return base.TableStart(t, row);
            
        }
        public override void TableCellStart(TableCell t, Row row)
        {
            base.TableCellStart(t, row);

            _currentWorkSheetSetting.NextCol();

            if (t.ColSpan>1)
            {
               _currentWorkSheetSetting.MergeCells.Add(string.Format("{0}{1}:{2}{3}"
                                            , (char)('A'+ _currentWorkSheetSetting.CurrentCol)
                                            ,_currentWorkSheetSetting.CurrentRow+1
                                            , (char)('A' + _currentWorkSheetSetting.CurrentCol + t.ColSpan-1 )
                                            , _currentWorkSheetSetting.CurrentRow+1 ));
            }
             
           
        }

        public override void TableCellEnd(TableCell t, Row row)
        {
            base.TableCellEnd(t, row);
            _currentWorkSheetSetting.NextCol(t.ColSpan - 1);
        }

        public override void TableRowStart(TableRow tr, Row row)
        {
            base.TableRowStart(tr, row);
            _currentWorkSheetSetting.NextRow();
            _currentWorkSheetSetting.StartCol();


        }
        public override void TableRowEnd(TableRow tr, Row row)
        {
            base.TableRowEnd(tr, row);

        }
        
        #endregion

        #region private methods
       private StyleInfo GetStyle(Style style, Row row)
        {
            if (style == null)
                return null;

            return style.GetStyleInfo(base.Report(), row);
        }
        private void GetStyledValue(Style style,Row row,ExcelRange range, string value)
        {
            StyleInfo pt = GetStyle(style, row);

            
            range.Style.Font.Bold = pt.IsFontBold();
            range.Style.Font.Italic = pt.FontStyle == FontStyleEnum.Italic;
            range.Style.Font.Size = pt.FontSize;

            if (pt.BStyleLeft != BorderStyleEnum.None)
            {
                range.Style.Border.Left.Style = GetBorderStyle(pt.BStyleLeft);
            }
            if (pt.BStyleRight != BorderStyleEnum.None)
            {
                range.Style.Border.Right.Style =GetBorderStyle(pt.BStyleRight);
            }
            if (pt.BStyleTop != BorderStyleEnum.None)
            {
                range.Style.Border.Top.Style =GetBorderStyle(pt.BStyleTop) ;
            }
            if (pt.BStyleBottom != BorderStyleEnum.None)
            {
                range.Style.Border.Bottom.Style = GetBorderStyle(pt.BStyleBottom) ;
            }

            
            
            range.Value = NumericValue(value) ?? value;
            if (IsNumeric(range.Value.ToString()))
                range.Style.Numberformat.Format = pt._Format;
            
            


        }

        private string NumericValue(string value)
        {

            value = value.Replace("(", "-");
            value = value.Replace(")", "");
            value = value.Replace(CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator, "");
            value = value.Replace(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator, ".");
            value = value.Replace("$", "");

            if (value.IndexOf('%') != -1)       //WRP 31102008 if a properly RDL formatted percentage need to remove "%" and pass throught value/100 to excel for correct formatting
            {
                value = value.Replace("%", String.Empty);
                decimal decvalue = Convert.ToDecimal(value) / 100;
                value = decvalue.ToString();
            }
            value = Regex.Replace(value, @"\s+", "");      //WRP 31102008 remove any white space
            return IsNumeric(value) ? value : null;
        }
        
        string GetColor(System.Drawing.Color c)
        {
            return GetColor("color", c);
        }

        string GetColor(string name, System.Drawing.Color c)
        {
            string s = string.Format("FF{0}{1}{2}", GetColor(c.R), GetColor(c.G), GetColor(c.B));
            return s;
        }

        string GetColor(byte b)
        {
            string sb = Convert.ToString(b, 16).ToUpperInvariant();
            return sb.Length > 1 ? sb : "0" + sb;
        }

        ExcelBorderStyle GetBorderStyle(BorderStyleEnum bs)
        {
            ExcelBorderStyle s;
            switch (bs)
            {
                case BorderStyleEnum.Solid:
                    s = ExcelBorderStyle.Thin;
                    break;
                case BorderStyleEnum.Dashed:
                    s = ExcelBorderStyle.Dashed;
                    break;
                case BorderStyleEnum.Dotted:
                    s = ExcelBorderStyle.Dotted;
                    break;
                case BorderStyleEnum.Double:
                    s = ExcelBorderStyle.Double;
                    break;
                case BorderStyleEnum.None:
                    s = ExcelBorderStyle.None;
                    break;
                default:
                    s = ExcelBorderStyle.Thin;
                    break;
            }
            return s;
        }


        #endregion



    }

    
}