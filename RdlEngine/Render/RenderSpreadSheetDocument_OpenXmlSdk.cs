
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

namespace fyiReporting.RDL
{


    ///<summary>
    /// Renders a report to PDF.   This is a page oriented formatting renderer.
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

        #endregion

        #region properties


        #endregion

        #region ctor
        public RenderSpreadSheetDocument_OpenXmlSdk(Report report, IStreamGen sg) : base(report, sg)
        {
            _ms = new MemoryStream();
            _spreadSheet = SpreadsheetDocument.Create(_ms, SpreadsheetDocumentType.Workbook);
            _openXmlExportHelper = new OpenXmlWriterHelper();
            _workbook = new Workbook();

        }
        #endregion

        #region implementations
        protected internal override void CreateDocument()
        {
            Report r = base.Report();
            _workbookPart = _spreadSheet.AddWorkbookPart();

            var openXmlExportHelper = new OpenXmlWriterHelper();
            openXmlExportHelper.SaveCustomStylesheet(_workbookPart);


            _workbookPart.Workbook = _workbook;
            var sheets = _workbook.AppendChild<Sheets>(new Sheets());



            // create worksheet 1
            var worksheetPart =_workbookPart.AddNewPart<WorksheetPart>();
            var sheet = new Sheet() { Id = _workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = r.Name };
            sheets.Append(sheet);
            _writer = OpenXmlWriter.Create(worksheetPart);
            _writer.WriteStartElement(new Worksheet());
            _writer.WriteStartElement(new SheetData());


        }

        protected internal override void EndDocument(Stream sg)
        {
            _writer.Close();
            //create the share string part using sax like approach too
            _openXmlExportHelper.CreateShareStringPart(_workbookPart);
            _spreadSheet.Close();

            byte[] contentbyte = _ms.ToArray();
            sg.Write(contentbyte, 0, contentbyte.Length);
            _ms.Dispose();
       
        }

        protected internal override void CreatePage()
        {
        }

        protected internal override void AfterProcessPage()
        {

        }

        protected internal override void AddBookmark(PageText pt)
        {

        }

        protected internal override void AddLine(float x, float y, float x2, float y2, float width, System.Drawing.Color c, BorderStyleEnum ls)
        {
        }

        protected internal override void AddImage(string name, StyleInfo si, ImageFormat imf, float x, float y, float width, float height, RectangleF clipRect, byte[] im, int samplesW, int samplesH, string url, string tooltip)
        {
        }

        protected internal override void AddPolygon(PointF[] pts, StyleInfo si, string url)
        {
        
        }

        protected internal override void AddRectangle(float x, float y, float height, float width, StyleInfo si, string url, string tooltip)
        {
        }


        protected internal override void AddPie(float x, float y, float height, float width, StyleInfo si, string url, string tooltip)
        {
        }

        protected internal override void AddCurve(PointF[] pts, StyleInfo si)
        {
        }
        protected internal override void AddEllipse(float x, float y, float height, float width, StyleInfo si, string url)
        {
        }
        protected internal override void AddText(PageText pt, Pages pgs)
        {
            
            if (_currentRow == null)
            {
                _currentRow =pt.Y;
                _writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row());
            }
            if (_currentRow !=pt.Y )
            {
                _currentRow = pt.Y;
                _writer.WriteEndElement();
                _writer.WriteStartElement(new DocumentFormat.OpenXml.Spreadsheet.Row());
            }
            
            _openXmlExportHelper.WriteCellValueSax(_writer,pt.Text , CellValues.InlineString);

           
        }

        #endregion

        #region private methods
        #endregion



    }
}