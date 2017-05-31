using System;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeWrapper
{
    public class PowerPoint
    {
        #region GlobalVars
        PPT.Application objApp;
        PPT.Presentations objPresSet;
        PPT._Presentation objPres;
        PPT.Slides objSlides;
        PPT._Slide objSlide;
        PPT.TextRange objTextRng;
        PPT.SlideShowTransition objSST;
        PPT.SlideRange objSldRng;
        PPT.Shapes objShapes;
        PPT.Shape objShape;
        PPT.Table objTable;
        PPT.Chart objPPTChart;
        PPT.ChartData objChartData;

        Excel.Workbook objWorkbook;
        Excel.Worksheet objSheet;

        int _shape;

        Color globalFontColor = Color.FromArgb(47, 86, 38);
        #endregion

        /// <summary>
        /// Create PowerPoint presentation
        /// </summary>
        public void createPPT(string templatePath = "", bool showPresentationWhileProcessing = false)
        {
            if (templatePath == "") templatePath = "\\\\bd-gsc-san2\\GSCDepts\\IT Dept\\DATA\\.net Assemblies\\BarrattTemplate.potx";

            // Must be visible if using charts
            MsoTriState visible;

            if (showPresentationWhileProcessing)
            {
                visible = MsoTriState.msoTrue;
            }
            else
            {
                visible = MsoTriState.msoFalse;
            }

            //Create a new presentation based on a template.
            objApp = new PPT.Application();
            //objApp.Visible = MsoTriState.msoFalse;
            objPresSet = objApp.Presentations;
            objPres = objPresSet.Open(templatePath, MsoTriState.msoFalse, MsoTriState.msoTrue, visible);
            objSlides = objPres.Slides;

            //Delete Template Slide
            objSlide = objSlides._Index(1);
            objSlide.Delete();
        }

        /// <summary>
        /// Adds Blank slide
        /// </summary>
        /// <param name="slideIndex"></param>
        public void addBlankSlide(int slideIndex)
        {
            _shape = 1;
            objSlide = objSlides.Add(slideIndex, PPT.PpSlideLayout.ppLayoutBlank);
        }

        /// <summary>
        /// Adds Blank slide
        /// </summary>
        /// <param name="slideIndex"></param>
        public void addTitleOnlySlide(int slideIndex)
        {
            _shape = 1;
            objSlide = objSlides.Add(slideIndex, PPT.PpSlideLayout.ppLayoutTitleOnly);
        }

        /// <summary>
        /// Set Title Slide Title
        /// </summary>
        /// <param name="text"></param>
        public void setTitleOnlySlideTitle(string text)
        {
            _shape = 1;
            objTextRng = objSlide.Shapes[_shape].TextFrame.TextRange;
            objTextRng.Text = text;
        }

        /// <summary>
        /// Table slide
        /// </summary>
        /// <param name="slideIndex"></param>
        /// <param name="slideTitle"></param>
        /// <param name="tableRows"></param>
        /// <param name="tableColumns"></param>
        public void AddTableSlide(int slideIndex, int tableRows, int tableColumns, string tableWidth = "960", string tableHeight = "540")
        {
            _shape = 1;
            if (tableWidth == "")
            {
                tableWidth = "960";
            }
            if (tableHeight == "")
            {
                tableHeight = "540";
            }

            objSlide = objSlides.Add(slideIndex, PPT.PpSlideLayout.ppLayoutBlank);
            //objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            //objTextRng.Text = slideTitle;
            //objTextRng.Font.Size = 48;

            objShapes = objSlide.Shapes;
            objShape = objShapes.AddTable(tableRows, tableColumns, 0, 0, float.Parse(tableWidth), float.Parse(tableHeight));
            objTable = objSlide.Shapes[1].Table;

            for (int col = 1; col <= tableColumns; col++)
            {
                for (int row = 1; row <= tableRows; row++)
                {
                    objTable.Cell(row, col).Shape.TextFrame.TextRange.Font.Size = 8;
                }
            }
        }

        /// <summary>
        /// Adds table
        /// </summary>
        /// <param name="slideIndex"></param>
        /// <param name="tableRows"></param>
        /// <param name="tableColumns"></param>
        /// <param name="tableWidth"></param>
        /// <param name="tableHeight"></param>
        /// <param name="tablePosFromTop"></param>
        /// <param name="tablePosFromLeft"></param>
        public void addTable(int tableColumns, int tableRows = 1, string tableWidth = "960", string tableHeight = "540", string tablePosFromTop = "0", string tablePosFromLeft = "0")
        {
            if (tableWidth == "")
            {
                tableWidth = "-1";
            }
            if (tableHeight == "")
            {
                tableHeight = "-1";
            }
            if (tablePosFromTop == "")
            {
                tablePosFromTop = "-1";
            }
            if (tablePosFromLeft == "")
            {
                tablePosFromLeft = "-1";
            }
            if (tableRows < 1)
            {
                tableRows = 1;
            }

            objShapes = objSlide.Shapes;
            objShape = objShapes.AddTable(tableRows, tableColumns, float.Parse(tablePosFromLeft), float.Parse(tablePosFromTop), float.Parse(tableWidth), float.Parse(tableHeight));
            _shape = objShape.Id;
            objTable = objShape.Table;
        }

        /// <summary>
        /// Adds cell content
        /// </summary>
        /// <param name="cellRow"></param>
        /// <param name="cellColumn"></param>
        /// <param name="cellText"></param>
        /// <param name="textBold"></param>
        /// <param name="textItalic"></param>
        /// <param name="textFont"></param>
        /// <param name="anchorVertical"></param>
        /// <param name="centerHorizontal"></param>
        public void addTableContent(int cellRow, int cellColumn, string cellText, bool textBold = false, bool textItalic = false, string textFont = "9", string anchorVertical = "top", bool centerHorizontal = false, bool textRotate = false, bool bulleted = false, string textColorCode = "", string cellColorCode = "")
        {
            if (textFont == "")
            {
                textFont = "9";
            }

            if (objTable.Rows.Count < cellRow)
            {
                objTable.Rows.Add().Height = 1;
            }

            cellText = cellText.Replace("^", Environment.NewLine);

            objTable.Cell(cellRow, cellColumn).Shape.TextFrame.TextRange.Text = cellText;
            if (textBold)
            {
                objTable.Cell(cellRow, cellColumn).Shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            }
            if (textItalic)
            {
                objTable.Cell(cellRow, cellColumn).Shape.TextFrame.TextRange.Font.Italic = MsoTriState.msoTrue;
            }
            if (bulleted)
            {
                objTable.Cell(cellRow, cellColumn).Shape.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = MsoTriState.msoTrue;
                objTable.Cell(cellRow, cellColumn).Shape.TextFrame.TextRange.ParagraphFormat.Bullet.Type = PPT.PpBulletType.ppBulletUnnumbered;
            }
            float size = float.Parse(textFont);

            objTable.Cell(cellRow, cellColumn).Shape.TextFrame.TextRange.Font.Size = size;

            if (textColorCode != "") objTable.Cell(cellRow, cellColumn).Shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml(textColorCode));

            switch (anchorVertical)
            {
                case "top":
                    objTable.Cell(cellRow, cellColumn).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
                    break;
                case "middle":
                    objTable.Cell(cellRow, cellColumn).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    break;
                case "bottom":
                    objTable.Cell(cellRow, cellColumn).Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                    break;
            }

            if (centerHorizontal)
            {
                objTable.Cell(cellRow, cellColumn).Shape.TextFrame.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorCenter;
            }

            if (cellColorCode != "") objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml(cellColorCode));

            if (textRotate) objTable.Cell(cellRow, cellColumn).Shape.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationUpward;

        }

        /// <summary>
        /// Sets the color of a cell
        /// </summary>
        /// <param name="cellRow"></param>
        /// <param name="cellColumn"></param>
        /// <param name="cellColor"></param>
        public void colorTableCell(int cellRow, int cellColumn, string cellColor)
        {
            //objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#00FF00"));
            switch (cellColor)
            {
                case "dark-green":
                    objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(47, 86, 38));
                    break;
                case "green":
                    objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(146, 208, 80));
                    break;
                case "amber":
                    objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(241, 149, 6));
                    break;
                case "red":
                    objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(251, 61, 50));
                    break;
                case "grey":
                    objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#969696"));
                    break;
                case "yellow":
                    objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFFF00"));
                    break;
                case "white":
                    objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFFFFF"));
                    break;
                default:
                    objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml(cellColor));
                    break;
            }
        }

        /// <summary>
        /// Resizes column
        /// </summary>
        /// <param name="column"></param>
        /// <param name="columnWidth"></param>
        public void resizeTableColumn(int column, string columnWidth)
        {
            objTable.Columns[column].Width = float.Parse(columnWidth);
        }

        /// <summary>
        /// Resizes row
        /// </summary>
        /// <param name="row"></param>
        /// <param name="rowHeight"></param>
        public void resizeTableRow(int row, string rowHeight)
        {
            objTable.Rows[row].Height = float.Parse(rowHeight);
        }

        /// <summary>
        /// Merges Cells in table
        /// </summary>
        /// <param name="cellRowFrom"></param>
        /// <param name="cellColumnFrom"></param>
        /// <param name="cellRowTo"></param>
        /// <param name="cellColumnTo"></param>
        public void mergeTableCells(int cellRowFrom, int cellColumnFrom, int cellRowTo, int cellColumnTo)
        {
            objTable.Cell(cellRowFrom, cellColumnFrom).Merge(objTable.Cell(cellRowTo, cellColumnTo));
        }

        /// <summary>
        /// Get X pos of cell
        /// </summary>
        /// <param name="cellRow"></param>
        /// <param name="cellColumn"></param>
        /// <returns></returns>
        public double getCellX(int cellRow, int cellColumn)
        {
            return objTable.Cell(cellRow, cellColumn).Shape.Left;
        }

        /// <summary>
        /// Get Y pos of cell
        /// </summary>
        /// <param name="cellRow"></param>
        /// <param name="cellColumn"></param>
        /// <returns></returns>
        public double getCellY(int cellRow, int cellColumn)
        {
            return objTable.Cell(cellRow, cellColumn).Shape.Top;
        }

        /// <summary>
        /// Get Height of cell
        /// </summary>
        /// <param name="cellRow"></param>
        /// <param name="cellColumn"></param>
        /// <returns></returns>
        public double getCellHeight(int cellRow, int cellColumn)
        {
            return objTable.Cell(cellRow, cellColumn).Shape.Height;
        }

        /// <summary>
        /// Get Width of cell
        /// </summary>
        /// <param name="cellRow"></param>
        /// <param name="cellColumn"></param>
        /// <returns></returns>
        public double getCellWidth(int cellRow, int cellColumn)
        {
            return objTable.Cell(cellRow, cellColumn).Shape.Width;
        }

        /// <summary>
        /// Sets styling
        /// </summary>
        public void styleTable(int style)
        {
            switch (style)
            {
                case 1:
                    objTable.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}", true);//No Style, Table Grid
                    for (int row = 1; row <= objTable.Rows.Count; row++)
                    {
                        for (int col = 1; col <= objTable.Columns.Count; col++)
                        {
                            objTable.Cell(row, col).Borders[PPT.PpBorderType.ppBorderLeft].Weight = 0.5f;
                            objTable.Cell(row, col).Borders[PPT.PpBorderType.ppBorderRight].Weight = 0.5f;
                            objTable.Cell(row, col).Borders[PPT.PpBorderType.ppBorderTop].Weight = 0.5f;
                            objTable.Cell(row, col).Borders[PPT.PpBorderType.ppBorderBottom].Weight = 0.5f;
                        }
                    }
                    break;
                case 2:
                    objTable.ApplyStyle("{912C8C85-51F0-491E-9774-3900AFEF0FD7}", true);//Light Style 2 - Accent 6
                    break;
                case 3:
                    objTable.ApplyStyle("{2A488322-F2BA-4B5B-9748-0D474271808F}", true);//Medium Style 3 - Accent 6
                    break;
                case 4:
                    objTable.ApplyStyle("{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}", true);//Dark Style 2 - Accent 5/Accent 6
                    break;
            }

        }

        /// <summary>
        /// Slide with title and subtitle
        /// </summary>
        /// <param name="slideIndex"></param>
        /// <param name="slideTitle"></param>
        /// <param name="slideSubTitle"></param>
        public void addTitleSlide(int slideIndex, string slideTitle, string slideSubTitle = "")
        {
            _shape = 1;
            objSlide = objSlides.Add(slideIndex, PPT.PpSlideLayout.ppLayoutTitle);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = slideTitle;
            //objTextRng.Font.Size = 60;

            objTextRng = objSlide.Shapes[2].TextFrame.TextRange;
            objTextRng.Text = slideSubTitle;
            //objTextRng.Font.Size = 24;
        }

        /// <summary>
        /// Slide with section title and subtitle
        /// </summary>
        /// <param name="slideIndex"></param>
        /// <param name="slideTitle"></param>
        /// <param name="slideSubTitle"></param>
        public void addSectionSlide(int slideIndex, string slideTitle, string slideSubTitle = "")
        {
            _shape = 1;
            objSlide = objSlides.Add(slideIndex, PPT.PpSlideLayout.ppLayoutSectionHeader);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = slideTitle;

            objTextRng = objSlide.Shapes[2].TextFrame.TextRange;
            objTextRng.Text = slideSubTitle;
        }

        /// <summary>
        /// Slide with title and text
        /// </summary>
        /// <param name="slideIndex"></param>
        /// <param name="slideTitle"></param>
        /// <param name="slideText"></param>
        public void addTextSlide(int slideIndex, string slideTitle)
        {
            _shape = 1;
            objSlide = objSlides.Add(slideIndex, PPT.PpSlideLayout.ppLayoutText);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = slideTitle;
            objTextRng.Font.Size = 48;
        }

        /// <summary>
        /// Paragraphs for Text Slide
        /// </summary>
        /// <param name="paragraphText"></param>
        public void addTextSlideParagraph(string paragraphText)
        {
            objTextRng = objSlide.Shapes[2].TextFrame.TextRange;
            if (objTextRng.Text == "")
            {
                objTextRng.Text = paragraphText;
            }
            else
            {
                objTextRng.Text = objTextRng.Text + "\n" + paragraphText;
            }
        }

        /// <summary>
        /// Slide with title and picture
        /// </summary>
        /// <param name="slideIndex"></param>
        /// <param name="slideTitle"></param>
        /// <param name="pictureFilename"></param>
        /// <param name="picturePosLeft"></param>
        /// <param name="picturePosTop"></param>
        /// <param name="picturetWidth"></param>
        /// <param name="pictureHeight"></param>
        public void addPictureSlide(int slideIndex, string slideTitle, string pictureFilename, int picturePosLeft, int picturePosTop, int pictureWidth, int pictureHeight)
        {
            _shape = 1;
            objSlide = objSlides.Add(slideIndex, PPT.PpSlideLayout.ppLayoutTitleOnly);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = slideTitle;
            objTextRng.Font.Size = 48;
            objSlide.Shapes.AddPicture(pictureFilename, MsoTriState.msoFalse, MsoTriState.msoTrue, picturePosLeft, picturePosTop, pictureWidth, pictureHeight);
        }

        /// <summary>
        /// Adds Text Box
        /// </summary>
        /// <param name="textboxWidth"></param>
        /// <param name="textboxHeight"></param>
        /// <param name="textboxPosFromTop"></param>
        /// <param name="textboxPosFromLeft"></param>
        /// <param name="textboxText"></param>
        public void addTextBox(string textboxWidth, string textboxHeight, string textboxPosFromTop, string textboxPosFromLeft)
        {
            objShapes = objSlide.Shapes;
            objShape = objShapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, float.Parse(textboxPosFromLeft), float.Parse(textboxPosFromTop), float.Parse(textboxWidth), float.Parse(textboxHeight));
            _shape++;
            objTextRng = objSlide.Shapes[_shape].TextFrame.TextRange;
        }

        /// <summary>
        /// Adds Paragraph to Text Box
        /// </summary>
        /// <param name="paragraphText"></param>
        public void addTextBoxParagraph(string paragraphText, bool bold = false, bool italic = false, string fontSize = "11", bool uppercase = false, bool useThemeFont = false)
        {
            if (fontSize == "")
            {
                fontSize = "11";
            }

            if (useThemeFont) objTextRng.Font.Color.RGB = globalFontColor.ToArgb(); objTextRng.Font.Name = "Arial";
            objTextRng.Font.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            objTextRng.Font.Italic = italic ? MsoTriState.msoTrue : MsoTriState.msoFalse;

            if (uppercase) objTextRng.ChangeCase(PPT.PpChangeCase.ppCaseUpper);

            objTextRng.Font.Size = float.Parse(fontSize);

            if (objTextRng.Text == "")
            {
                objTextRng.InsertAfter(paragraphText);
            }
            else
            {
                objTextRng.InsertAfter("\n" + paragraphText);
            }

        }

        /// <summary>
        /// Adds a Shape
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="shapeName"></param>
        /// <param name="fillColorHex"></param>
        /// <param name="text"></param>
        /// <param name="textColorHex"></param>
        /// <param name="textFontSize"></param>
        /// <param name="borderColorHex"></param>
        public void addShape(string x, string y, string width, string height, string shapeName = "rectangle", string fillColorHex = "", string text = "", string textColorHex = "", string textFontSize = "0", string borderColorHex = "")
        {
            if (float.Parse(width) <= 0) width = "0";

            objShapes = objSlide.Shapes;
            MsoAutoShapeType msoShapeType;
            switch (shapeName)
            {
                case "oval":
                    msoShapeType = MsoAutoShapeType.msoShapeOval;
                    break;
                case "diamond":
                    msoShapeType = MsoAutoShapeType.msoShapeDiamond;
                    break;
                case "arrow":
                    msoShapeType = MsoAutoShapeType.msoShapeLeftRightArrow;
                    break;
                default:
                    msoShapeType = MsoAutoShapeType.msoShapeRectangle;
                    break;
            }
            objShape = objShapes.AddShape(msoShapeType, float.Parse(x), float.Parse(y), float.Parse(width), float.Parse(height));
            _shape++;
            if (fillColorHex != "") objShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml(fillColorHex));
            if (borderColorHex != "")
            {
                //objShape.Line.Style = MsoLineStyle.msoLineThinThin;
                objShape.Line.Weight = 1.5f;
                objShape.Line.BackColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml(borderColorHex));
            }
            else
            {
                objShape.Line.Visible = MsoTriState.msoFalse;
            }
            if (text != "")
            {
                objTextRng = objShape.TextFrame.TextRange;
                objTextRng.Text = text;
                if (textColorHex != "") objTextRng.Font.Color.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml(textColorHex));
                if (textFontSize != "0") objTextRng.Font.Size = float.Parse(textFontSize);
            }
        }

        /// <summary>
        /// Sets the slide to specified index
        /// </summary>
        /// <param name="slideIndex"></param>
        public void setCurrentSlide(int slideIndex)
        {
            objSlide = objSlides[slideIndex];
        }

        /// <summary>
        /// Gets shape's height
        /// </summary>
        /// <returns></returns>
        public int getCurrentShapeHeight()
        {
            return (int)objShape.Height;
        }

        /// <summary>
        /// Gets shape's width
        /// </summary>
        /// <returns></returns>
        public int getCurrentShapeWidth()
        {
            return (int)objShape.Width;
        }

        /// <summary>
        /// Get presentation slide's height
        /// </summary>
        /// <returns></returns>
        public int getSlideHeight()
        {
            return (int)objPres.PageSetup.SlideHeight;
        }

        /// <summary>
        /// Get presentation slide's width
        /// </summary>
        /// <returns></returns>
        public int getSlideWidth()
        {
            return (int)objPres.PageSetup.SlideWidth;
        }

        /// <summary>
        /// Import another presentation
        /// </summary>
        /// <param name="slideIndex"></param>
        /// <param name="filepath"></param>
        /// <returns></returns>
        public int importPresentation(int slideIndex, string filepath)
        {
            int slideID = objSlides.InsertFromFile(filepath, slideIndex);
            objSlide = objSlides[slideID];
            return slideID;
        }

        /// <summary>
        /// Step 4 - Set transitions for each slide
        /// </summary>
        public void setTransisions(string transitionType = "fade")
        {
            PPT.PpEntryEffect entryEffect;
            switch (transitionType)
            {
                case "fade":
                    entryEffect = PPT.PpEntryEffect.ppEffectFade;
                    break;
                case "fadesmooth":
                    entryEffect = PPT.PpEntryEffect.ppEffectFadeSmoothly;
                    break;
                case "pandown":
                    entryEffect = PPT.PpEntryEffect.ppEffectPanDown;
                    break;
                case "panup":
                    entryEffect = PPT.PpEntryEffect.ppEffectPanUp;
                    break;
                case "panleft":
                    entryEffect = PPT.PpEntryEffect.ppEffectPanLeft;
                    break;
                case "panright":
                    entryEffect = PPT.PpEntryEffect.ppEffectPanRight;
                    break;
                case "random":
                    entryEffect = PPT.PpEntryEffect.ppEffectRandom;
                    break;
                default:
                    entryEffect = PPT.PpEntryEffect.ppEffectFade;
                    break;
            }

            //Modify the slide show transition settings for all slides in the presentation.
            //int[] SlideIdx = new int[slideAmount];
            //for (int i = 0; i < slideAmount; i++)
            //{
            //    SlideIdx[i] = i + 1;
            //}
            objSldRng = objSlides.Range(objSlides.Count);
            objSST = objSldRng.SlideShowTransition;
            //objSST.AdvanceOnTime = MsoTriState.msoTrue;
            //objSST.AdvanceTime = 3;
            objSST.EntryEffect = entryEffect;

        }

        /// <summary>
        /// Save to filename
        /// </summary>
        /// <param name="filename"></param>
        public void savePPT(string filename, bool showPPT = false)
        {
            string path = filename;
            System.Threading.Thread.Sleep(1000);
            try
            {
                string filenm = path.Substring(path.LastIndexOf("\\") + 1);
                path = path.Substring(0, path.LastIndexOf("\\") + 1);
                filenm = filenm.Replace("<", " ")
                    .Replace(">", " ")
                    .Replace(":", " ")
                    .Replace("\"", " ")
                    .Replace("/", " ")
                    .Replace("|", " ")
                    .Replace("?", " ")
                    .Replace("*", " ");
                filename = path + filenm;

                objPres.SaveAs(filename);
                if (showPPT) objPres.Close(); objPres = objPresSet.Open(filename, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue); objApp.Activate(); objApp.WindowState = PPT.PpWindowState.ppWindowMinimized; objApp.WindowState = PPT.PpWindowState.ppWindowMaximized;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                MessageBox.Show(filename + " May be currently in use. Please exit any instances of it and click OK. \n\n Additional Details: \n" + e.Message);
                savePPT(filename);
            }
        }

        /// <summary>
        /// Show Presentation
        /// </summary>
        public void showPPT(string filename)
        {
            if (objPres != null) objPres.Close();
            objPres = objPresSet.Open(filename, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            objApp.Activate();
            objApp.WindowState = PPT.PpWindowState.ppWindowMinimized;
            objApp.WindowState = PPT.PpWindowState.ppWindowMaximized;
        }

        /// <summary>
        /// Debugger from string
        /// </summary>
        /// <param name="text"></param>
        public void debugText(string text)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            using (System.IO.StreamWriter outputFile = new System.IO.StreamWriter(desktopPath + @"\OfficeWrapperDebug.txt", true))
            {
                outputFile.WriteLine(text);
            }
        }

        #region Experimental
        //// WORKING BUT ONLY IF PPT IS OPEN - HIDDEN FOR NOW

        /// <summary>
        /// Setup and create chart slide WORK IN PROGRESS
        /// </summary>
        /// <param name="slideIndex"></param>
        /// <param name="cellFrom"></param>
        /// <param name="cellTo"></param>
        /// <param name="chartPosLeft"></param>
        /// <param name="chartPosTop"></param>
        /// <param name="chartWidth"></param>
        /// <param name="chartHeight"></param>
        /// <param name="chartType"></param>
        /// <param name="chartLegendPosition"></param>
        private void AddChartSlide(int slideIndex, bool chartFillsSlide = true)
        {
            _shape = 1;

            //int chartWidth, int chartHeight, int chartPosLeft = -1, int chartPosTop = -1

            float chartPosLeft;
            float chartPosTop;
            float chartWidth;
            float chartHeight;

            if (chartFillsSlide)
            {
                chartPosLeft = 0;
                chartPosTop = 0;
                chartWidth = 960;
                chartHeight = 540;
            }
            else
            {
                chartPosLeft = -1;
                chartPosTop = -1;
                chartWidth = -1;
                chartHeight = -1;
            }

            objSlide = objSlides.Add(slideIndex, PPT.PpSlideLayout.ppLayoutBlank);

            objPPTChart = objSlide.Shapes.AddChart(XlChartType.xlLine, chartPosLeft, chartPosTop, chartWidth, chartHeight).Chart;

            objChartData = objPPTChart.ChartData;

            objWorkbook = (Excel.Workbook)objChartData.Workbook;

            objSheet = (Excel.Worksheet)objWorkbook.Worksheets[1];


            //Clear Template
            objSheet.Cells.ClearContents();
        }

        /// <summary>
        /// Add chart data to worksheet
        /// </summary>
        /// <param name="range"></param>
        /// <param name="value"></param>
        private void AddChartData(string range, string value)
        {
            objSheet.Cells.get_Range(range).FormulaR1C1 = value;
        }

        /// <summary>
        /// Format chart
        /// </summary>
        /// <param name="chartHasTitle"></param>
        /// <param name="chartTitle"></param>
        private void FormatChart(bool chartHasTitle, bool chartHasLabels, string chartTitle = "", string categoryTitle = "", string ValueTitle = "", bool dataIsLabeled = false, string chartType = "line", string chartLegendPosition = "R", string resizeChartAreaTo = "")
        {
            Excel.Range last = objSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range tRange = objSheet.Cells.get_Range("A1", last);

            if (resizeChartAreaTo != "")
            {
                tRange = objSheet.Cells.get_Range("A1", resizeChartAreaTo);
            }

            int lastUsedRow = last.Row;
            int lastUsedColumn = last.Column;

            objSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "Table1";
            tRange.Select();
            objSheet.ListObjects["Table1"].TableStyle = "TableStyleMedium15";
            objSheet.ListObjects["Table1"].Resize(tRange);




            if (chartHasTitle)
            {
                objPPTChart.HasTitle = chartHasTitle;
                objPPTChart.ChartTitle.Text = chartTitle;
            }

            switch (chartType)
            {
                case "line":
                    objPPTChart.ChartType = XlChartType.xlLine;
                    break;
                case "column":
                    objPPTChart.ChartType = XlChartType.xlColumnClustered;
                    break;
                case "bar":
                    objPPTChart.ChartType = XlChartType.xlBarClustered;
                    break;
                case "pie":
                    objPPTChart.ChartType = XlChartType.xl3DPie;
                    break;
                case "area":
                    objPPTChart.ChartType = XlChartType.xlArea;
                    break;
                case "scatter":
                    objPPTChart.ChartType = XlChartType.xlXYScatter;
                    break;
                default:
                    objPPTChart.ChartType = XlChartType.xlLine;
                    break;
            }

            switch (chartLegendPosition)
            {
                case "T":
                    objPPTChart.Legend.Position = PPT.XlLegendPosition.xlLegendPositionTop;
                    break;
                case "B":
                    objPPTChart.Legend.Position = PPT.XlLegendPosition.xlLegendPositionBottom;
                    break;
                case "L":
                    objPPTChart.Legend.Position = PPT.XlLegendPosition.xlLegendPositionLeft;
                    break;
                case "R":
                    objPPTChart.Legend.Position = PPT.XlLegendPosition.xlLegendPositionRight;
                    break;
                default:
                    objPPTChart.Legend.Position = PPT.XlLegendPosition.xlLegendPositionRight;
                    break;
            }


            //objPPTChart.ChartTitle.Text = "2007 Sales";
            //objPPTChart.ChartTitle.Font.Italic = true;
            //objPPTChart.ChartTitle.Font.Size = 18;
            //objPPTChart.ChartTitle.Font.Color = Color.Black.ToArgb();
            //objPPTChart.ChartTitle.Format.Line.Visible = MsoTriState.msoTrue;
            //objPPTChart.ChartTitle.Format.Line.ForeColor.RGB = Color.Black.ToArgb();

            if (chartHasLabels)
            {
                objPPTChart.ChartWizard
                    (Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, categoryTitle, ValueTitle);
                //objPPTChart.Axes(PPT.XlAxisType.xlCategory, PPT.XlAxisGroup.xlPrimary);
            }
            if (dataIsLabeled)
            {
                objPPTChart.ApplyDataLabels(PPT.XlDataLabelsType.xlDataLabelsShowLabel);
            }
        }

        /// <summary>
        /// Close workbook
        /// </summary>
        private void CloseWorkbook()
        {
            objWorkbook.Application.Quit();
            //objWorkbook.Close();
        }
        #endregion

    }
}