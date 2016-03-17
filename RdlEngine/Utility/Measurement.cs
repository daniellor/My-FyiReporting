/*
 * This file is part of reportFU, based on the work of 
 * Kim Sheffield and the fyiReporting project. 
 *
 * Copyright (c) 2010 devFU Pty Ltd, Josh Wilson and Others (http://reportfu.org)
 * 
 * Prior Copyrights:
 * _________________________________________________________
 * |Copyright (C) 2004-2008  fyiReporting Software, LLC     |
 * |For additional information, email info@fyireporting.com |
 * |or visit the website www.fyiReporting.com.              |
 * =========================================================
 *
 * License:
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace fyiReporting.RDL.Utility
{
    /// <summary>
    /// A utility class that contains additional 
    /// methods for drawing and unit conversion (points/pixels).
    /// </summary>
    public sealed class Measurement
    {

        static readonly char[] LINEBREAK = new char[] { '\n' };
        static readonly char[] WORDBREAK = new char[] { ' ' };
        //		static readonly int MEASUREMAX = int.MaxValue;  //  .Net 2 doesn't seem to have a limit; 1.1 limit was 32
        static readonly int MEASUREMAX = 32;  //  guess I'm wrong -- .Net 2 doesn't seem to have a limit; 1.1 limit was 32

        struct WordStartFinish
        {
            internal float start;
            internal float end;
        }


        /// <summary>
        /// A method used to obtain a rectangle from the screen coordinates supplied.
        /// </summary>
        public static System.Drawing.Rectangle RectFromPoints(Point p1, Point p2)
        {
            System.Drawing.Rectangle r = new System.Drawing.Rectangle();
            // set the width and x of rectangle
            if (p1.X < p2.X)
            {
                r.X = p1.X;
                r.Width = p2.X - p1.X;
            }
            else
            {
                r.X = p2.X;
                r.Width = p1.X - p2.X;
            }
            // set the height and y of rectangle
            if (p1.Y < p2.Y)
            {
                r.Y = p1.Y;
                r.Height = p2.Y - p1.Y;
            }
            else
            {
                r.Y = p2.Y;
                r.Height = p1.Y - p2.Y;
            }
            return r;
        }

        /// <summary>
        /// The constant value used to calculate the conversion of pixels into points, as a float.
        /// </summary>
        public const float POINTSIZE_F = 72.27f;
        /// <summary>
        /// The constant value used to calculate the conversion of pixels into points, as a decimal.
        /// </summary>
        public const decimal POINTSIZE_M = 72.27m;

        public const float STANDARD_DPI_X = 96f;
        public const float STANDARD_DPI_Y = 96f;

        /// <summary>
        /// A method used to convert pixels into points.
        /// </summary>
        /// <returns>A float containing the converted measurement of the pixels into points.</returns>
        public static float PointsFromPixels(float pixels, float dpi)
        {
            return (pixels * POINTSIZE_F) / dpi;
        }
        /// <summary>
        /// A method used to convert pixels into points.
        /// </summary>
        /// <returns>A PointF containing the point X and Y values for the pixel X and Y values that were supplied.</returns>
        public static PointF PointsFromPixels(float pixelsX, float pixelsY, PointF Dpi)
        {
            return new PointF(PointsFromPixels(pixelsX, Dpi.X), PointsFromPixels(pixelsY, Dpi.Y));
        }
        /// <summary>
        /// A method used to convert points into pixels.
        /// </summary>
        /// <returns>An int containing the converted measurement of the points into pixels.</returns>
        public static int PixelsFromPoints(float points, float dpi)
        {
            int r = (int)(((double)points * dpi) / POINTSIZE_F);
            if (r == 0 && points > .0001f)
                r = 1;
            return r;
        }
        /// <summary>
        /// A method used to convert points into pixels.
        /// </summary>
        /// <returns>A PointF containing the pixel X and Y values for the point X and Y values that were supplied.</returns>
        public static PointF PixelsFromPoints(float pointsX, float pointsY, PointF Dpi)
        {
            return new PointF(PixelsFromPoints(pointsX, Dpi.X), PixelsFromPoints(pointsY, Dpi.Y));
        }
        /// <summary>
        /// A method used to convert points into twips.
        /// </summary>
        /// <returns>An int containing the twips for the number of points that were supplied.</returns>
        public static int TwipsFromPoints(float points)
        {
            return (int)Math.Round(points * 20, 0);
        }
        /// <summary>
        /// A method used to convert pixels into twips.
        /// </summary>
        /// <returns>An int containing the twips for the number of pixels that were supplied.</returns>
        public static int TwipsFromPixels(float pixels, float dpi)
        {
            return TwipsFromPoints(PointsFromPixels(pixels, dpi));
        }


        public static string[] MeasureString(PageText pt, Graphics g, out float[] width)
        {
            StyleInfo si = pt.SI;
            string s = pt.Text;

            System.Drawing.Font drawFont = null;
            StringFormat drawFormat = null;
            SizeF ms;
            string[] sa = null;
            width = null;
            try
            {
                // STYLE
                System.Drawing.FontStyle fs = 0;
                if (si.FontStyle == FontStyleEnum.Italic)
                    fs |= System.Drawing.FontStyle.Italic;

                // WEIGHT
                switch (si.FontWeight)
                {
                    case FontWeightEnum.Bold:
                    case FontWeightEnum.Bolder:
                    case FontWeightEnum.W500:
                    case FontWeightEnum.W600:
                    case FontWeightEnum.W700:
                    case FontWeightEnum.W800:
                    case FontWeightEnum.W900:
                        fs |= System.Drawing.FontStyle.Bold;
                        break;
                    default:
                        break;
                }

                drawFont = new System.Drawing.Font(StyleInfo.GetFontFamily(si.FontFamilyFull), si.FontSize, fs);
                drawFormat = new StringFormat();
                drawFormat.Alignment = StringAlignment.Near;

                // Measure string   
                //  pt.NoClip indicates that this was generated by PageTextHtml Build.  It has already word wrapped.
                if (pt.NoClip || pt.SI.WritingMode == WritingModeEnum.tb_rl)	// TODO: support multiple lines for vertical text
                {
                    ms = MeasureString(s, g, drawFont, drawFormat);
                    width = new float[1];
                    width[0] = Measurement.PointsFromPixels(ms.Width, g.DpiX);	// convert to points from pixels
                    sa = new string[1];
                    sa[0] = s;
                    return sa;
                }

                // handle multiple lines;
                //  1) split the string into the forced line breaks (ie "\n and \r")
                //  2) foreach of the forced line breaks; break these into words and recombine 
                s = s.Replace("\r\n", "\n");	// don't want this to result in double lines
                string[] flines = s.Split(LINEBREAK);
                List<string> lines = new List<string>();
                List<float> lineWidths = new List<float>();
                // remove the size reserved for left and right padding
                float ptWidth = pt.W - pt.SI.PaddingLeft - pt.SI.PaddingRight;
                if (ptWidth <= 0)
                    ptWidth = 1;
                foreach (string tfl in flines)
                {
                    string fl;
                    if (tfl.Length > 0 && tfl[tfl.Length - 1] == ' ')
                        fl = tfl.TrimEnd(' ');
                    else
                        fl = tfl;

                    // Check if entire string fits into a line
                    ms = MeasureString(fl, g, drawFont, drawFormat);
                    float tw = Measurement.PointsFromPixels(ms.Width, g.DpiX);
                    if (tw <= ptWidth)
                    {					   // line fits don't need to break it down further
                        lines.Add(fl);
                        lineWidths.Add(tw);
                        continue;
                    }

                    // Line too long; need to break into multiple lines
                    // 1) break line into parts; then build up again keeping track of word positions
                    string[] parts = fl.Split(WORDBREAK);	// this is the maximum split of lines
                    StringBuilder sb = new StringBuilder(fl.Length);
                    CharacterRange[] cra = new CharacterRange[parts.Length];
                    for (int i = 0; i < parts.Length; i++)
                    {
                        int sc = sb.Length;	 // starting character
                        sb.Append(parts[i]);	// endding character
                        if (i != parts.Length - 1)  // last item doesn't need blank
                            sb.Append(" ");
                        int ec = sb.Length;
                        CharacterRange cr = new CharacterRange(sc, ec - sc);
                        cra[i] = cr;			// add to character array
                    }

                    // 2) Measure the word locations within the line
                    string wfl = sb.ToString();
                    WordStartFinish[] wordLocations = MeasureString(wfl, g, drawFont, drawFormat, cra);
                    if (wordLocations == null)
                        continue;

                    // 3) Loop thru creating new lines as needed
                    int startLoc = 0;
                    CharacterRange crs = cra[startLoc];
                    CharacterRange cre = cra[startLoc];
                    float cwidth = wordLocations[0].end;	// length of the first
                    float bwidth = wordLocations[0].start;  // characters need a little extra on start
                    string ts;
                    bool bLine = true;
                    for (int i = 1; i < cra.Length; i++)
                    {
                        cwidth = wordLocations[i].end - wordLocations[startLoc].start + bwidth;
                        if (cwidth > ptWidth)
                        {	// time for a new line
                            cre = cra[i - 1];
                            ts = wfl.Substring(crs.First, cre.First + cre.Length - crs.First);
                            lines.Add(ts);
                            lineWidths.Add(wordLocations[i - 1].end - wordLocations[startLoc].start + bwidth);

                            // Find the first non-blank character of the next line
                            while (i < cra.Length &&
                                    cra[i].Length == 1 &&
                                    fl[cra[i].First] == ' ')
                            {
                                i++;
                            }
                            if (i < cra.Length)   // any lines left?
                            {  // yes, continue on
                                startLoc = i;
                                crs = cre = cra[startLoc];
                                cwidth = wordLocations[i].end - wordLocations[startLoc].start + bwidth;
                            }
                            else  // no, we can stop
                                bLine = false;
                            //  bwidth = wordLocations[startLoc].start - wordLocations[startLoc - 1].end;
                        }
                        else
                            cre = cra[i];
                    }
                    if (bLine)
                    {
                        ts = fl.Substring(crs.First, cre.First + cre.Length - crs.First);
                        lines.Add(ts);
                        lineWidths.Add(cwidth);
                    }
                }
                // create the final array from the Lists
                string[] la = lines.ToArray();
                width = lineWidths.ToArray();
                return la;
            }
            finally
            {
                if (drawFont != null)
                    drawFont.Dispose();
                if (drawFormat != null)
                    drawFont.Dispose();
            }
        }

        private static SizeF MeasureString(string s, Graphics g, System.Drawing.Font drawFont, StringFormat drawFormat)
        {
            if (s == null || s.Length == 0)
                return SizeF.Empty;

            CharacterRange[] cr = { new CharacterRange(0, s.Length) };
            drawFormat.SetMeasurableCharacterRanges(cr);
            Region[] rs = new Region[1];
            rs = g.MeasureCharacterRanges(s, drawFont, new RectangleF(0, 0, float.MaxValue, float.MaxValue),
                drawFormat);
            RectangleF mr = rs[0].GetBounds(g);

            return new SizeF(mr.Width, mr.Height);
        }
        /// <summary>
        /// Measures the location of an arbritrary # of words within a string
        /// </summary>
        private static WordStartFinish[] MeasureString(string s, Graphics g, System.Drawing.Font drawFont, StringFormat drawFormat, CharacterRange[] cra)
        {
            if (cra.Length <= MEASUREMAX)		// handle the simple case of < MEASUREMAX words
                return MeasureString32(s, g, drawFont, drawFormat, cra);

            // Need to compensate for SetMeasurableCharacterRanges limitation of 32 (MEASUREMAX)
            int mcra = (cra.Length / MEASUREMAX);	// # of full 32 arrays we need
            int ip = cra.Length % MEASUREMAX;		// # of partial entries needed for last array (if any)
            WordStartFinish[] sz = new WordStartFinish[cra.Length];	// this is the final result;
            float startPos = 0;
            CharacterRange[] cra32 = new CharacterRange[MEASUREMAX];	// fill out			
            int icra = 0;						// index thru the cra 
            for (int i = 0; i < mcra; i++)
            {
                // fill out the new array
                int ticra = icra;
                for (int j = 0; j < cra32.Length; j++)
                {
                    cra32[j] = cra[ticra++];
                    cra32[j].First -= cra[icra].First;	// adjust relative offsets of strings
                }

                // measure the word locations (in the new string)
                // ???? should I put a blank in front of it?? 
                string ts = s.Substring(cra[icra].First,
                    cra[icra + cra32.Length - 1].First + cra[icra + cra32.Length - 1].Length - cra[icra].First);
                WordStartFinish[] pos = MeasureString32(ts, g, drawFont, drawFormat, cra32);

                // copy the values adding in the new starting positions
                for (int j = 0; j < pos.Length; j++)
                {
                    sz[icra].start = pos[j].start + startPos;
                    sz[icra++].end = pos[j].end + startPos;
                }
                startPos = sz[icra - 1].end;	// reset the start position for the next line
            }
            // handle the remaining character
            if (ip > 0)
            {
                // resize the range array
                cra32 = new CharacterRange[ip];
                // fill out the new array
                int ticra = icra;
                for (int j = 0; j < cra32.Length; j++)
                {
                    cra32[j] = cra[ticra++];
                    cra32[j].First -= cra[icra].First;	// adjust relative offsets of strings
                }
                // measure the word locations (in the new string)
                // ???? should I put a blank in front of it?? 
                string ts = s.Substring(cra[icra].First,
                    cra[icra + cra32.Length - 1].First + cra[icra + cra32.Length - 1].Length - cra[icra].First);
                WordStartFinish[] pos = MeasureString32(ts, g, drawFont, drawFormat, cra32);

                // copy the values adding in the new starting positions
                for (int j = 0; j < pos.Length; j++)
                {
                    sz[icra].start = pos[j].start + startPos;
                    sz[icra++].end = pos[j].end + startPos;
                }
            }
            return sz;
        }

        /// <summary>
        /// Measures the location of words within a string;  limited by .Net 1.1 to 32 words
        ///	 MEASUREMAX is a constant that defines that limit
        /// </summary>
        /// <param name="s"></param>
        /// <param name="g"></param>
        /// <param name="drawFont"></param>
        /// <param name="drawFormat"></param>
        /// <param name="cra"></param>
        /// <returns></returns>
        private static WordStartFinish[] MeasureString32(string s, Graphics g, System.Drawing.Font drawFont, StringFormat drawFormat, CharacterRange[] cra)
        {
            if (s == null || s.Length == 0)
                return null;

            drawFormat.SetMeasurableCharacterRanges(cra);
            Region[] rs = new Region[cra.Length];
            rs = g.MeasureCharacterRanges(s, drawFont, new RectangleF(0, 0, float.MaxValue, float.MaxValue),
                drawFormat);
            WordStartFinish[] sz = new WordStartFinish[cra.Length];
            int isz = 0;
            foreach (Region r in rs)
            {
                RectangleF mr = r.GetBounds(g);
                sz[isz].start = Measurement.PointsFromPixels(mr.Left, g.DpiX);
                sz[isz].end = Measurement.PointsFromPixels(mr.Right, g.DpiX);
                isz++;
            }
            return sz;
        }

        // convert points to Excel units: characters 
        //   Assume 11 characters to the inch
        public static float PointsToExcelUnits(float pointWidth)
        {
            return (float)(pointWidth / POINTSIZE_F) * 11; 
        }


        #region Obsolete Methods
        /// <summary>
        /// A method used to convert Pixels into Points. Obsolete. Use PointsFromPixels instead.
        /// </summary>
        /// <returns>A float containing the converted measurement of the pixels into points.</returns>
        [System.Obsolete("This method has been deprecated. Use PointsFromPixels() instead.")]
        public static float PointsX(float pixelX, float dpi)// pixels to points
        {
            return PointsFromPixels(pixelX, dpi);
        }
        /// <summary>
        /// A method used to convert Pixels into Points. Obsolete. Use PointsFromPixels instead.
        /// </summary>
        /// <returns>A float containing the converted measurement of the pixels into points.</returns>
        [System.Obsolete("This method has been deprecated. Use PointsFromPixels() instead.")]
        public static float PointsY(float pixelY, float dpi)
        {
            return PointsFromPixels(pixelY, dpi);
        }

        /// <summary>
        /// A method used to convert Points into Pixels. Obsolete. Use PixelsFromPoints instead.
        /// </summary>
        /// <returns>An int containing the converted measurement of the points into pixels.</returns>
        [System.Obsolete("This method has been deprecated. Use PixelsFromPoints() instead.")]
        public static int PixelsX(float pointX, float dpi)// points to pixels
        {
            return PixelsFromPoints(pointX, dpi);
        }


        /// <summary>
        /// A method used to convert Points into Pixels. Obsolete. Use PixelsFromPoints instead.
        /// </summary>
        /// <returns>An int containing the converted measurement of the points into pixels.</returns>
        [System.Obsolete("This method has been deprecated. Use PixelsFromPoints() instead.")]
        public static int PixelsY(float pointY, float dpi)
        {
            return PixelsFromPoints(pointY, dpi);
        }
        #endregion
    }
}