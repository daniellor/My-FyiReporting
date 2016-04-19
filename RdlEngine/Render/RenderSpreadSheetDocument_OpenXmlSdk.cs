
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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;

namespace fyiReporting.RDL
{


    ///<summary>
    /// Renders a report to SpreadSheet.
    ///</summary>
    [SecuritySafeCritical]
    internal class RenderSpreadSheetDocument_OpenXmlSdk : RenderBase
    {
        #region private
        SpreadsheetDocument _spreadSheet;
        MemoryStream _ms;
        OpenXmlWriter _writer;
        float? _currentRow = null;
        OpenXmlWriterHelper _openXmlExportHelper ;
        Workbook _workbook; 
        WorkbookPart _workbookPart ;
        Stylesheet _styleSheet;
        List<WorkSheetSetting> _workSheetSettings;

        WorkSheetSetting _currentWorkSheet;
        #endregion

        

        #region ctor
        public RenderSpreadSheetDocument_OpenXmlSdk(Report report, IStreamGen sg) : base(report, sg)
        {
            _ms = new MemoryStream();
            _spreadSheet = SpreadsheetDocument.Create(_ms, SpreadsheetDocumentType.Workbook);
            _openXmlExportHelper = new OpenXmlWriterHelper();
            _workbook = new Workbook();
            _workSheetSettings = new List<WorkSheetSetting>();

        }
        #endregion

        #region implementation abstract methods
        protected internal override void CreateDocument()
        {
            Report r = base.Report();
            _workbookPart = _spreadSheet.AddWorkbookPart();

            var openXmlExportHelper = new OpenXmlWriterHelper();
            _styleSheet = openXmlExportHelper.CreateDefaultStylesheet();

            _workbookPart.Workbook = _workbook;
            var sheets = _workbook.AppendChild<Sheets>(new Sheets());



            // create worksheet 1
            var worksheetPart =_workbookPart.AddNewPart<WorksheetPart>();
            var sheet = new Sheet() { Id = _workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = r.Name };
            sheets.Append(sheet);


            _workSheetSettings.Add(new WorkSheetSetting(r.Name));

            _writer = OpenXmlWriter.Create(worksheetPart);
            _writer.WriteStartElement(new Worksheet());
            _writer.WriteStartElement(new SheetData());

           _currentWorkSheet = _workSheetSettings.Last();
            

        }

      
        protected internal override void EndDocument(Stream sg)
        {
         
            var workbookStylesPart = _workbookPart.AddNewPart<WorkbookStylesPart>();
            var style = workbookStylesPart.Stylesheet = _styleSheet;
            style.Save();

            _writer.Close();
            //create the share string part using sax like approach too
            _openXmlExportHelper.CreateShareStringPart(_workbookPart);

            

           foreach (var worksheetSetting in _workSheetSettings)
            {
                Worksheet worksheet = _openXmlExportHelper.GetWorksheet(_spreadSheet, worksheetSetting.WorksheetName);
                foreach (var mergeCell in worksheetSetting.MergeCells)
                     _openXmlExportHelper.MergeCells(worksheet, mergeCell);
                
            }
            _spreadSheet.Close();

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
                if (_currentRow == null || _currentRow != tb.Top.PixelsY)
                {
                        _currentRow = tb.Top.PixelsY;
                        _writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row()); //new row
                }
                
            }

            int? idCellFormat = GetStyleIndex(tb.Style, row, ref t);

            if (idCellFormat != null)
            {

                var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, idCellFormat.ToString()) }.ToList();
              //  if (tb.Width!=null)
              //      attributes.Add(new OpenXmlAttribute("width", null, Measurement.PointsToExcelUnits(tb.Width.Points).ToString()));
                if ((IsNumeric(t) || IsNumeric(t, CultureInfo.CurrentCulture)))
                    _openXmlExportHelper.WriteCellValueSax(_writer, t, CellValues.Number, attributes);
                else
                    _openXmlExportHelper.WriteCellValueSax(_writer, t, CellValues.InlineString, attributes);
            }
            else
                _openXmlExportHelper.WriteCellValueSax(_writer, t, CellValues.InlineString);

            if (!tableCell)
            {
                _writer.WriteEndElement();
                _currentWorkSheet.NextRow();
            }
           


        }

        public override bool TableStart(Table t, Row row)
        {
            return base.TableStart(t, row);
            
        }
        public override void TableCellStart(TableCell t, Row row)
        {
            base.TableCellStart(t, row);

            _currentWorkSheet.NextCol();

            if (t.ColSpan>1)
            {
               _currentWorkSheet.MergeCells.Add(string.Format("{0}{1}:{2}{3}"
                                            , (char)('A'+ _currentWorkSheet.CurrentCol)
                                            ,_currentWorkSheet.CurrentRow+1
                                            , (char)('A' + _currentWorkSheet.CurrentCol + t.ColSpan-1 )
                                            , _currentWorkSheet.CurrentRow+1 ));
            }
             
           
        }

        public override void TableCellEnd(TableCell t, Row row)
        {
            base.TableCellEnd(t, row);
            _currentWorkSheet.NextCol(t.ColSpan - 1);
        }

        public override void TableRowStart(TableRow tr, Row row)
        {
            base.TableRowStart(tr, row);
            _currentWorkSheet.NextRow();
            _currentWorkSheet.StartCol();
            _writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row());


        }
        public override void TableRowEnd(TableRow tr, Row row)
        {
            base.TableRowEnd(tr, row);
            _writer.WriteEndElement();

        }
        
        #endregion

        #region private methods
       private StyleInfo GetStyle(Style style, Row row)
        {
            if (style == null)
                return null;

            return style.GetStyleInfo(base.Report(), row);
        }
        private int? GetStyleIndex(Style style,Row row, ref string value)
        {
            StyleInfo pt = GetStyle(style, row);
            // DocumentFormat.OpenXml.Spreadsheet.Fonts fonts1 = new DocumentFormat.OpenXml.Spreadsheet.Fonts()
            //                                                     { Count = (UInt32Value)1U, KnownFonts = true };
            int? fontId = null;
            int? borderId = null;
            int? cellFormatId = null;

            DocumentFormat.OpenXml.Spreadsheet.Font font = new DocumentFormat.OpenXml.Spreadsheet.Font();

            if (pt.IsFontBold())
                font.Append(new DocumentFormat.OpenXml.Spreadsheet.Bold());
            if (pt.FontStyle == FontStyleEnum.Italic)
                font.Append(new DocumentFormat.OpenXml.Spreadsheet.Italic());

            font.Append(new DocumentFormat.OpenXml.Spreadsheet.FontSize() { Val = (Double)pt.FontSize });
            font.Append(new DocumentFormat.OpenXml.Spreadsheet.FontName() { Val = pt.FontFamily });
            //font.Append(new DocumentFormat.OpenXml.Spreadsheet.Color()
            //            {  Rgb=GetColor(si.Color) });


            int id = 0;
            foreach (var fo in _styleSheet.Fonts)
            {
                if (fo.OuterXml.Equals(font.OuterXml))
                {
                    fontId = id;
                    break;
                }
                id++;
            }

            if (fontId == null)
            {
                _styleSheet.Fonts.Append(font);
                _styleSheet.Fonts.Count = (uint)_styleSheet.Fonts.ChildElements.Count;
                fontId = _styleSheet.Fonts.ChildElements.Count - 1;
            }

            Border border = new Border();
            if (pt.BStyleLeft != BorderStyleEnum.None)
            {
                border.LeftBorder = new LeftBorder() { Style = GetBorderStyle(pt.BStyleLeft) };
            }
            if (pt.BStyleRight != BorderStyleEnum.None)
            {
                border.RightBorder = new RightBorder() { Style = GetBorderStyle(pt.BStyleRight) };
            }
            if (pt.BStyleTop != BorderStyleEnum.None)
            {
                border.TopBorder = new TopBorder() { Style = GetBorderStyle(pt.BStyleTop) };
            }
            if (pt.BStyleBottom != BorderStyleEnum.None)
            {
                border.BottomBorder = new BottomBorder() { Style = GetBorderStyle(pt.BStyleBottom) };
            }

            id = 0;
            foreach (var bo in _styleSheet.Borders)
            {
                if (bo.OuterXml.Equals(border.OuterXml))
                {
                    borderId = id;
                    break;
                }
                id++;
            }

            if (borderId == null)
            {
                _styleSheet.Borders.Append(border);
                _styleSheet.Borders.Count = (uint)_styleSheet.Borders.ChildElements.Count;
                borderId = _styleSheet.Borders.ChildElements.Count - 1;
            }


            value = NumericValue(value) ?? value;


            CellFormat cf = new CellFormat();
            cf.NumberFormatId = (uint)StyleInfo.GetFormatCode(pt._Format);
            cf.FontId = (uint)fontId;
            cf.FillId = 0;
            cf.BorderId = (uint)borderId;
            // cf.FormatId = 0;


            id = 0;
            foreach (var cef in _styleSheet.CellFormats)
            {
                if (cef.OuterXml.Equals(cf.OuterXml))
                {
                    cellFormatId = id;
                    break;
                }
                id++;
            }

            if (cellFormatId == null)
            {
                _styleSheet.CellFormats.Append(cf);
                _styleSheet.CellFormats.Count = (uint)_styleSheet.CellFormats.ChildElements.Count;
                cellFormatId = _styleSheet.CellFormats.ChildElements.Count - 1;
            }


            return cellFormatId;



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

        BorderStyleValues GetBorderStyle(BorderStyleEnum bs)
        {
            BorderStyleValues s;
            switch (bs)
            {
                case BorderStyleEnum.Solid:
                    s = BorderStyleValues.Thin;
                    break;
                case BorderStyleEnum.Dashed:
                    s = BorderStyleValues.Dashed;
                    break;
                case BorderStyleEnum.Dotted:
                    s = BorderStyleValues.Dotted;
                    break;
                case BorderStyleEnum.Double:
                    s = BorderStyleValues.Double;
                    break;
                case BorderStyleEnum.None:
                    s = BorderStyleValues.None;
                    break;
                default:
                    s = BorderStyleValues.Thin;
                    break;
            }
            return s;
        }


        #endregion



    }

    internal class WorkSheetSetting
    {
        public string WorksheetName;
        public int CurrentRow { get; private set; }
        public int CurrentCol { get; private set; }

        public List<string> MergeCells;

        public WorkSheetSetting(string worksheetName)
        {
            WorksheetName = worksheetName;
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
            return CurrentCol+increment;
        }
        public int StartCol()
        {
            return CurrentCol= -1;
        }
        public int StartRow()
        {
            return CurrentRow = -1;
        }
        public int NextRow()
        {
            return CurrentRow++;
        }
        public int NextRow(int increment)
        {
            return CurrentRow+increment;
        }


    }
}