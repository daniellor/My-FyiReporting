
/*
 * 
 Copyright (C) 2004-2008  fyiReporting Software, LLC
 Copyright (C) 2011  Peter Gill <peter@majorsilence.com>
 Copyright (c) 2010 devFU Pty Ltd, Josh Wilson and Others (http://reportfu.org)



 This file has been modified with suggestiong from forum users.
 *Obtained from Forum, User: sinnovasoft http://www.fyireporting.com/forum/viewtopic.php?t=1049

 Refactored by Daniel Romanowski http://dotlink.pl

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
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using fyiReporting.RDL.Utility;
using System.Security;

namespace fyiReporting.RDL
{


    ///<summary>
    /// Renders a report to PDF.   This is a page oriented formatting renderer.
    ///</summary>
    [SecuritySafeCritical]
    internal abstract class RenderBase : IPresent
    {

        #region private
        Stream _streamGen;                  // where the output is going
        Report _report;                 // report
    
        #endregion


        #region properties
        private PdfPageSize _pageSize;

        internal protected PdfPageSize PageSize
        {
            get { return _pageSize; }
            private set { _pageSize = value; }
        }
        #endregion


        #region abstract methods
        internal protected void  AddLine(float x, float y, float x2, float y2, StyleInfo si)
        {
            AddLine(x, y, x2, y2, si.BWidthTop, si.BColorTop, si.BStyleTop);
        }
        /// <summary>
        /// Page line element at the X Y to X2 Y2 position
        /// </summary>
        /// <returns></returns>
        internal abstract protected void CreateDocument();
        internal abstract protected void EndDocument(Stream sg);
        internal abstract protected void CreatePage();
        internal abstract protected void AfterProcessPage();
        internal abstract protected void AddBookmark(PageText pt);

        internal abstract protected void AddLine(float x, float y, float x2, float y2, float width, System.Drawing.Color c, BorderStyleEnum ls);
       
      
        /// <summary>
        /// Add image to the page.
        /// </summary>
        /// <returns>string Image name</returns>
        internal abstract protected void AddImage(string name,  StyleInfo si,
            ImageFormat imf, float x, float y, float width, float height, RectangleF clipRect,
            byte[] im, int samplesW, int samplesH, string url, string tooltip);

        /// <summary>
        /// Page Polygon
        /// </summary>
        /// <param name="pts"></param>
        /// <param name="si"></param>
        /// <param name="url"></param>
        /// <param name="patterns"></param>
        internal abstract protected void AddPolygon(PointF[] pts, StyleInfo si, string url);

      
        /// <summary>
        /// Page Rectangle element at the X Y position
        /// </summary>
        /// <returns></returns>
        internal abstract protected void AddRectangle(float x, float y, float height, float width, StyleInfo si, string url,  string tooltip);
        /// <summary>
        /// Draw a pie
        /// </summary>
        /// <returns></returns>
        internal abstract protected void AddPie(float x, float y, float height, float width, StyleInfo si, string url,  string tooltip);

        /// <summary>
        /// Draw a curve
        /// </summary>
        /// <returns></returns>
        internal abstract protected void AddCurve(PointF[] pts, StyleInfo si);


      
        //25072008 GJL Draw 4 bezier curves to approximate a circle
        internal abstract protected void AddEllipse(float x, float y, float height, float width, StyleInfo si, string url);

        

        /// <summary>
        /// Page Text element at the X Y position; multiple lines handled
        /// </summary>
        /// <returns></returns>
        internal abstract protected void AddText(PageText pt, Pages pgs);

        #endregion

        //Replaced from forum, User: Aulofee http://www.fyireporting.com/forum/viewtopic.php?t=793
        public void Dispose() { }

      
        public RenderBase(Report rep, IStreamGen sg)
        {
            _streamGen = sg.GetStream();
            _report = rep;
        }

        public Report Report()
        {
            return _report;
        }

        public bool IsPagingNeeded()
        {
            return true;
        }

        public void Start()
        {
            CreateDocument();
        }

        public void End()
        {
            EndDocument(_streamGen);
            return;
        }

        public void RunPages(Pages pgs)	// this does all the work
        {
            foreach (Page p in pgs)
            {
                PageSize = new PdfPageSize((int)_report.ReportDefinition.PageWidth.ToPoints(),
                                       (int)_report.ReportDefinition.PageHeight.ToPoints());
             
                //Create a Page 
                CreatePage();               
                ProcessPage(pgs, p);
                // after a page
                AfterProcessPage();
            }
            return;
        }
        // render all the objects in a page in PDF
        private void ProcessPage(Pages pgs, IEnumerable items)
        {
            foreach (PageItem pi in items)
            {
                if (pi.SI.BackgroundImage != null)
                {	// put out any background image
                    PageImage bgImg = pi.SI.BackgroundImage;
                    //					elements.AddImage(images, i.Name, content.objectNum, i.SI, i.ImgFormat, 
                    //						pi.X, pi.Y, pi.W, pi.H, i.ImageData,i.SamplesW, i.SamplesH, null);				   
                    //Duc Phan modified 10 Dec, 2007 to support on background image 
                    float imW = Measurement.PointsFromPixels(bgImg.SamplesW, pgs.G.DpiX);
                    float imH = Measurement.PointsFromPixels(bgImg.SamplesH, pgs.G.DpiY);
                    int repeatX = 0;
                    int repeatY = 0;
                    float itemW = pi.W - (pi.SI.PaddingLeft + pi.SI.PaddingRight);
                    float itemH = pi.H - (pi.SI.PaddingTop + pi.SI.PaddingBottom);
                    switch (bgImg.Repeat)
                    {
                        case ImageRepeat.Repeat:
                            repeatX = (int)Math.Floor(itemW / imW);
                            repeatY = (int)Math.Floor(itemH / imH);
                            break;
                        case ImageRepeat.RepeatX:
                            repeatX = (int)Math.Floor(itemW / imW);
                            repeatY = 1;
                            break;
                        case ImageRepeat.RepeatY:
                            repeatY = (int)Math.Floor(itemH / imH);
                            repeatX = 1;
                            break;
                        case ImageRepeat.NoRepeat:
                        default:
                            repeatX = repeatY = 1;
                            break;
                    }

                    //make sure the image is drawn at least 1 times 
                    repeatX = Math.Max(repeatX, 1);
                    repeatY = Math.Max(repeatY, 1);

                    float currX = pi.X + pi.SI.PaddingLeft;
                    float currY = pi.Y + pi.SI.PaddingTop;
                    float startX = currX;
                    float startY = currY;
                    for (int i = 0; i < repeatX; i++)
                    {
                        for (int j = 0; j < repeatY; j++)
                        {
                            currX = startX + i * imW;
                            currY = startY + j * imH;
                       
                           

                                AddImage( bgImg.Name,bgImg.SI, bgImg.ImgFormat,
                                                currX, currY, imW, imH, RectangleF.Empty, bgImg.ImageData, bgImg.SamplesW, bgImg.SamplesH, null, pi.Tooltip);
                           
                        }
                    }
                }

                if (pi is PageTextHtml)
                {
                    PageTextHtml pth = pi as PageTextHtml;
                    pth.Build(pgs.G);
                    ProcessPage(pgs, pth);
                    continue;
                }

                if (pi is PageText)
                {
                    PageText pt = pi as PageText;
                    

                    AddText(pt,pgs);
                    
                    if (pt.Bookmark != null)
                    {
                        AddBookmark(pt);
                    }
                    continue;
                }

                if (pi is PageLine)
                {
                    PageLine pl = pi as PageLine;
                    AddLine(pl.X, pl.Y, pl.X2, pl.Y2, pl.SI);
                    continue;
                }

                if (pi is PageEllipse)
                {
                    PageEllipse pe = pi as PageEllipse;
                    AddEllipse(pe.X, pe.Y, pe.H, pe.W, pe.SI, pe.HyperLink);
                    continue;
                }



                if (pi is PageImage)
                {
                    PageImage i = pi as PageImage;

                    //Duc Phan added 20 Dec, 2007 to support sized image 
                    RectangleF r2 = new RectangleF(i.X + i.SI.PaddingLeft, i.Y + i.SI.PaddingTop, i.W - i.SI.PaddingLeft - i.SI.PaddingRight, i.H - i.SI.PaddingTop - i.SI.PaddingBottom);

                    RectangleF adjustedRect;   // work rectangle 
                    RectangleF clipRect = RectangleF.Empty;
                    switch (i.Sizing)
                    {
                        case ImageSizingEnum.AutoSize:
                            adjustedRect = new RectangleF(r2.Left, r2.Top,
                                            r2.Width, r2.Height);
                            break;
                        case ImageSizingEnum.Clip:
                            adjustedRect = new RectangleF(r2.Left, r2.Top,
                                            Measurement.PointsFromPixels(i.SamplesW, pgs.G.DpiX), Measurement.PointsFromPixels(i.SamplesH, pgs.G.DpiY));
                            clipRect = new RectangleF(r2.Left, r2.Top,
                                            r2.Width, r2.Height);
                            break;
                        case ImageSizingEnum.FitProportional:
                            float height;
                            float width;
                            float ratioIm = (float)i.SamplesH / i.SamplesW;
                            float ratioR = r2.Height / r2.Width;
                            height = r2.Height;
                            width = r2.Width;
                            if (ratioIm > ratioR)
                            {   // this means the rectangle width must be corrected 
                                width = height * (1 / ratioIm);
                            }
                            else if (ratioIm < ratioR)
                            {   // this means the rectangle height must be corrected 
                                height = width * ratioIm;
                            }
                            adjustedRect = new RectangleF(r2.X, r2.Y, width, height);
                            break;
                        case ImageSizingEnum.Fit:
                        default:
                            adjustedRect = r2;
                            break;
                    }
                    if (i.ImgFormat == System.Drawing.Imaging.ImageFormat.Wmf || i.ImgFormat == System.Drawing.Imaging.ImageFormat.Emf)
                    {
                        //We dont want to add it - its already been broken down into page items;
                    }
                    else
                    {
                       
                            AddImage(i.Name,  i.SI, i.ImgFormat,
                            adjustedRect.X, adjustedRect.Y, adjustedRect.Width, adjustedRect.Height, clipRect, i.ImageData, i.SamplesW, i.SamplesH, i.HyperLink, i.Tooltip);
                       
                    }
                    continue;
                }

                if (pi is PageRectangle)
                {
                    PageRectangle pr = pi as PageRectangle;
                    AddRectangle(pr.X, pr.Y, pr.H, pr.W, pi.SI, pi.HyperLink,  pi.Tooltip);
                    continue;
                }
                if (pi is PagePie)
                {   // TODO
                    PagePie pp = pi as PagePie;
                    // 
                    AddPie(pp.X, pp.Y, pp.H, pp.W, pi.SI, pi.HyperLink,  pi.Tooltip);
                    continue;
                }
                if (pi is PagePolygon)
                {
                    PagePolygon ppo = pi as PagePolygon;
                    AddPolygon(ppo.Points, pi.SI, pi.HyperLink);
                    continue;
                }
                if (pi is PageCurve)
                {
                    PageCurve pc = pi as PageCurve;
                    AddCurve(pc.Points, pi.SI);
                    continue;
                }

            }

        }

       
       
     
        // Body: main container for the report
        public void BodyStart(Body b)
        {
        }

        public void BodyEnd(Body b)
        {
        }

        public void PageHeaderStart(PageHeader ph)
        {
        }

        public void PageHeaderEnd(PageHeader ph)
        {
        }

        public void PageFooterStart(PageFooter pf)
        {
        }

        public void PageFooterEnd(PageFooter pf)
        {
        }

        public void Textbox(Textbox tb, string t, Row row)
        {
        }

        public void DataRegionNoRows(DataRegion d, string noRowsMsg)
        {
        }

        // Lists
        public bool ListStart(List l, Row r)
        {
            return true;
        }

        public void ListEnd(List l, Row r)
        {
        }

        public void ListEntryBegin(List l, Row r)
        {
        }

        public void ListEntryEnd(List l, Row r)
        {
        }

        // Tables					// Report item table
        public bool TableStart(Table t, Row row)
        {
            return true;
        }

        public void TableEnd(Table t, Row row)
        {
        }

        public void TableBodyStart(Table t, Row row)
        {
        }

        public void TableBodyEnd(Table t, Row row)
        {
        }

        public void TableFooterStart(Footer f, Row row)
        {
        }

        public void TableFooterEnd(Footer f, Row row)
        {
        }

        public void TableHeaderStart(Header h, Row row)
        {
        }

        public void TableHeaderEnd(Header h, Row row)
        {
        }

        public void TableRowStart(TableRow tr, Row row)
        {
        }

        public void TableRowEnd(TableRow tr, Row row)
        {
        }

        public void TableCellStart(TableCell t, Row row)
        {
            return;
        }

        public void TableCellEnd(TableCell t, Row row)
        {
            return;
        }

        public bool MatrixStart(Matrix m, MatrixCellEntry[,] matrix, Row r, int headerRows, int maxRows, int maxCols)				// called first
        {
            return true;
        }

        public void MatrixColumns(Matrix m, MatrixColumns mc)	// called just after MatrixStart
        {
        }

        public void MatrixCellStart(Matrix m, ReportItem ri, int row, int column, Row r, float h, float w, int colSpan)
        {
        }

        public void MatrixCellEnd(Matrix m, ReportItem ri, int row, int column, Row r)
        {
        }

        public void MatrixRowStart(Matrix m, int row, Row r)
        {
        }

        public void MatrixRowEnd(Matrix m, int row, Row r)
        {
        }

        public void MatrixEnd(Matrix m, Row r)				// called last
        {
        }

        public void Chart(Chart c, Row r, ChartBase cb)
        {
        }

        public void Image(fyiReporting.RDL.Image i, Row r, string mimeType, Stream ior)
        {
        }

        public void Line(Line l, Row r)
        {
            return;
        }

        public bool RectangleStart(fyiReporting.RDL.Rectangle rect, Row r)
        {
            return true;
        }

        public void RectangleEnd(fyiReporting.RDL.Rectangle rect, Row r)
        {
        }

        public void Subreport(Subreport s, Row r)
        {
        }

        public void GroupingStart(Grouping g)			// called at start of grouping
        {
        }
        public void GroupingInstanceStart(Grouping g)	// called at start for each grouping instance
        {
        }
        public void GroupingInstanceEnd(Grouping g)	// called at start for each grouping instance
        {
        }
        public void GroupingEnd(Grouping g)			// called at end of grouping
        {
        }
    }
}