using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeWrapper
{
    public class Document
    {
        #region Globals
        Word.Application wordApp;
        Word.Document wordDoc;
        Word.Table wordTable;
        #endregion

        /// <summary>
        /// Creates Word Document
        /// </summary>
        /// <param name="templatePath"></param>
        public void createWordDoc(string templatePath = "")
        {
            try
            {
                wordApp = new Word.Application();
                wordApp.ShowAnimation = false;
                wordApp.Visible = false;
                //Create a new document
                wordDoc = wordApp.Documents.Add(templatePath != "" ? templatePath : "", Missing.Value, Missing.Value, false);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }

        }

        /// <summary>
        /// Adds Paragraph to Doc
        /// </summary>
        /// <param name="text"></param>
        /// <param name="useStyleNo"></param>
        public void addParagraph(string text, string bookmarkName = "", string styleName = "", bool textBold = false, bool textItalic = false, string textFontSize = "", string horizontalAlignment = "")
        {
            Word.Range range = bookmarkName.Equals("")
                ? wordDoc.Range()
                : wordDoc.Bookmarks[bookmarkName].Range;
            text.Replace("\n", Environment.NewLine);
            range.Text = text;
            range.Font.Bold = textBold ? 1 : 0;
            range.Font.Italic = textItalic ? 1 : 0;
            if (textFontSize != "") range.Font.Size = float.Parse(textFontSize);
            switch (horizontalAlignment)
            {
                case "left":
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    break;
                case "center":
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    break;
                case "right":
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    break;
                case "justify":
                    range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    break;
            }
            if (styleName != "") range.set_Style(styleName);
        }

        /// <summary>
        /// Sets Header
        /// </summary>
        /// <param name="text"></param>
        public void setHeader(string text)
        {
            try
            {

                //Add header into the document
                foreach (Word.Section section in wordDoc.Sections)
                {
                    //Get the header range and add the header details.
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    //headerRange.Font.Size = 10;
                    //headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                    headerRange.Text = text;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        /// <summary>
        /// Sets footer
        /// </summary>
        /// <param name="text"></param>
        public void setFooter(string text)
        {
            try
            {

                //Add the footers into the document
                foreach (Word.Section wordSection in wordDoc.Sections)
                {
                    //Get the footer range and add the footer details.
                    Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                    //footerRange.Font.Size = 10;
                    //footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    footerRange.Text = text;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        //#region List
        //public void createList()
        //{
        //    OfficeList.createList();
        //}
        //
        //public void addToList(string col01, string col02 = "", string col03 = "", string col04 = "", string col05 = "",
        //    string col06 = "", string col07 = "", string col08 = "", string col09 = "", string col10 = "",
        //    string col11 = "", string col12 = "", string col13 = "", string col14 = "", string col15 = "",
        //    string col16 = "", string col17 = "", string col18 = "", string col19 = "", string col20 = "")
        //{
        //    OfficeList.addToList(new OfficeItem(col01, col02, col03, col04, col05, col06, col07, col08, col09, col10,
        //        col11, col12, col13, col14, col15, col16, col17, col18, col19, col20));
        //}
        //
        //public void addTableFromList(int columns, string bookmarkName = "")
        //{
        //    Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add(Missing.Value);
        //
        //    Word.Table table = wordDoc.Tables.Add(paragraph.Range, 1, columns, Missing.Value, Word.WdAutoFitBehavior.wdAutoFitContent);
        //    table.Borders.Enable = 1;
        //
        //    foreach (OfficeItem officeItem in OfficeList.officeList)
        //    {
        //        //for (int i = 0; i >= columns; i++)
        //        foreach (Word.Cell cell in table.Rows.Add().Cells)
        //        {
        //            switch (cell.ColumnIndex)
        //            {
        //                case 1:
        //                    cell.Range.Text = officeItem.col01;
        //                    break;
        //                case 2:
        //                    cell.Range.Text = officeItem.col02;
        //                    break;
        //                case 3:
        //                    cell.Range.Text = officeItem.col03;
        //                    break;
        //                case 4:
        //                    cell.Range.Text = officeItem.col04;
        //                    break;
        //                case 5:
        //                    cell.Range.Text = officeItem.col05;
        //                    break;
        //                case 6:
        //                    cell.Range.Text = officeItem.col06;
        //                    break;
        //                case 7:
        //                    cell.Range.Text = officeItem.col07;
        //                    break;
        //                case 8:
        //                    cell.Range.Text = officeItem.col08;
        //                    break;
        //                case 9:
        //                    cell.Range.Text = officeItem.col09;
        //                    break;
        //                case 10:
        //                    cell.Range.Text = officeItem.col10;
        //                    break;
        //                case 11:
        //                    cell.Range.Text = officeItem.col11;
        //                    break;
        //                case 12:
        //                    cell.Range.Text = officeItem.col12;
        //                    break;
        //                case 13:
        //                    cell.Range.Text = officeItem.col13;
        //                    break;
        //                case 14:
        //                    cell.Range.Text = officeItem.col14;
        //                    break;
        //                case 15:
        //                    cell.Range.Text = officeItem.col15;
        //                    break;
        //                case 16:
        //                    cell.Range.Text = officeItem.col16;
        //                    break;
        //                case 17:
        //                    cell.Range.Text = officeItem.col17;
        //                    break;
        //                case 18:
        //                    cell.Range.Text = officeItem.col18;
        //                    break;
        //                case 19:
        //                    cell.Range.Text = officeItem.col19;
        //                    break;
        //                case 20:
        //                    cell.Range.Text = officeItem.col20;
        //                    break;
        //            }
        //        }
        //    }
        //
        //    //firstTable.Borders.Enable = 1;
        //    //foreach (Word.Row row in table.Rows)
        //    //{
        //    //    foreach (Word.Cell cell in row.Cells)
        //    //    {
        //    //        //Header row
        //    //        if (cell.RowIndex == 1)
        //    //        {
        //    //            cell.Range.Text = "Column " + cell.ColumnIndex.ToString();
        //    //            cell.Range.Font.Bold = 1;
        //    //            //other format properties goes here
        //    //            cell.Range.Font.Name = "verdana";
        //    //            cell.Range.Font.Size = 10;
        //    //            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                            
        //    //            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
        //    //            //Center alignment for the Header cells
        //    //            cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        //    //            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        //    //
        //    //        }
        //    //        //Data row
        //    //        else
        //    //        {
        //    //            cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
        //    //        }
        //    //    }
        //    //}
        //}
        //#endregion

        /// <summary>
        /// Adds content to prexisting table with ID
        /// </summary>
        /// <param name="tableID"></param>
        /// <param name="cellRow"></param>
        /// <param name="cellColumn"></param>
        /// <param name="text"></param>
        public void addToExistingTable(string tableID, int cellRow, int cellColumn, string text)
        {
            int id = int.Parse(tableID);
            Word.Table table = wordDoc.Tables[id];
            Word.Cell cell = table.Cell(cellRow, cellColumn);
            cell.Range.InsertAfter(text);
        }

        /// <summary>
        /// Adds table
        /// </summary>
        /// <param name="tableColumns"></param>
        /// <param name="tableRows"></param>
        /// <param name="bookmarkName"></param>
        /// <param name="tableBorders"></param>
        public void addTable(int tableColumns, int tableRows = 1, string bookmarkName = "", bool tableBorders = false)
        {
            if (tableRows < 1)
            {
                tableRows = 1;
            }
            Word.Range wordRange = bookmarkName.Equals("")
                ? wordDoc.Range()
                : wordDoc.Bookmarks[bookmarkName].Range;

            wordTable = wordDoc.Tables.Add(wordRange, tableRows, tableColumns, Missing.Value, Word.WdAutoFitBehavior.wdAutoFitContent);
            wordTable.Borders.Enable = tableBorders ? 1 : 0;
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
        /// <param name="horizontalAlignment"></param>
        /// <param name="textColorCode"></param>
        public void addTableContent(int cellRow, int cellColumn, string cellText, bool textBold = false, bool textItalic = false, string textFontSize = "9", string anchorVertical = "top", string horizontalAlignment = "left", string textColorCode = "")
        {
            if (textFontSize == "")
            {
                textFontSize = "9";
            }

            if (wordTable.Rows.Count < cellRow)
            {
                wordTable.Rows.Add().Height = 1;
            }

            cellText = cellText.Replace("^", Environment.NewLine);

            wordTable.Cell(cellRow, cellColumn).Range.Text = cellText;
            wordTable.Cell(cellRow, cellColumn).Range.Font.Bold = textBold ? 1 : 0;
            wordTable.Cell(cellRow, cellColumn).Range.Font.Italic = textItalic ? 1 : 0;
            wordTable.Cell(cellRow, cellColumn).Range.Font.Size = float.Parse(textFontSize);

            wordTable.Cell(cellRow, cellColumn).Range.Font.TextColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml(textColorCode != "" ? textColorCode : "#000000"));

            switch (anchorVertical)
            {
                case "top":
                    wordTable.Cell(cellRow, cellColumn).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                    break;
                case "center":
                    wordTable.Cell(cellRow, cellColumn).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    break;
                case "bottom":
                    wordTable.Cell(cellRow, cellColumn).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                    break;
            }

            switch (horizontalAlignment)
            {
                case "left":
                    wordTable.Cell(cellRow, cellColumn).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    break;
                case "center":
                    wordTable.Cell(cellRow, cellColumn).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    break;
                case "right":
                    wordTable.Cell(cellRow, cellColumn).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    break;
                case "justify":
                    wordTable.Cell(cellRow, cellColumn).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    break;
            }


        }

        /// <summary>
        /// Sets the color of a cell
        /// </summary>
        /// <param name="cellRow"></param>
        /// <param name="cellColumn"></param>
        /// <param name="cellColor"></param>
        public void colorTableCell(int cellRow, int cellColumn, string cellColorCode)
        {
            wordTable.Cell(cellRow, cellColumn).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(ColorTranslator.FromHtml(cellColorCode));
            //switch (cellColor)
            //{
            //    case "dark-green":
            //        wordTable.Cell(cellRow, cellColumn).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(Color.FromArgb(47, 86, 38));
            //        break;
            //    case "green":
            //        //objTable.Cell(cellRow, cellColumn).Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(ColorTranslator.FromHtml("#00FF00"));
            //        wordTable.Cell(cellRow, cellColumn).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(Color.FromArgb(146, 208, 80));
            //        break;
            //    case "amber":
            //        wordTable.Cell(cellRow, cellColumn).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(Color.FromArgb(241, 149, 6));
            //        break;
            //    case "red":
            //        wordTable.Cell(cellRow, cellColumn).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(Color.FromArgb(251, 61, 50));
            //        break;
            //    case "grey":
            //        wordTable.Cell(cellRow, cellColumn).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(ColorTranslator.FromHtml("#969696"));
            //        break;
            //    case "yellow":
            //        wordTable.Cell(cellRow, cellColumn).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFFF00"));
            //        break;
            //    case "white":
            //        wordTable.Cell(cellRow, cellColumn).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFFFFF"));
            //        break;
            //}
        }

        /// <summary>
        /// Resizes column
        /// </summary>
        /// <param name="column"></param>
        /// <param name="columnWidth"></param>
        public void resizeTableColumn(int column, string columnWidth)
        {
            wordTable.Columns[column].Width = float.Parse(columnWidth);
        }

        /// <summary>
        /// Resizes row
        /// </summary>
        /// <param name="row"></param>
        /// <param name="rowHeight"></param>
        public void resizeTableRow(int row, string rowHeight)
        {
            wordTable.Rows[row].Height = float.Parse(rowHeight);
        }

        public void autoFitColumn(int column)
        {
            wordTable.Columns[column].AutoFit();
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
            wordTable.Cell(cellRowFrom, cellColumnFrom).Merge(wordTable.Cell(cellRowTo, cellColumnTo));
        }

        /// <summary>
        /// Sets styling
        /// </summary>
        //public void styleTable(int style)
        //{
        //switch (style)
        //{
        //    case 1:
        //        wordTable.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}", true);//No Style, Table Grid
        //        for (int row = 1; row <= wordTable.Rows.Count; row++)
        //        {
        //            for (int col = 1; col <= wordTable.Columns.Count; col++)
        //            {
        //                wordTable.Cell(row, col).Borders[PPT.PpBorderType.ppBorderLeft].Weight = 0.5f;
        //                wordTable.Cell(row, col).Borders[PPT.PpBorderType.ppBorderRight].Weight = 0.5f;
        //                wordTable.Cell(row, col).Borders[PPT.PpBorderType.ppBorderTop].Weight = 0.5f;
        //                wordTable.Cell(row, col).Borders[PPT.PpBorderType.ppBorderBottom].Weight = 0.5f;
        //            }
        //        }
        //        break;
        //    case 2:
        //        wordTable.ApplyStyle("{912C8C85-51F0-491E-9774-3900AFEF0FD7}", true);//Light Style 2 - Accent 6
        //        break;
        //    case 3:
        //        wordTable.ApplyStyle("{2A488322-F2BA-4B5B-9748-0D474271808F}", true);//Medium Style 3 - Accent 6
        //        break;
        //    case 4:
        //        wordTable.ApplyStyle("{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}", true);//Dark Style 2 - Accent 5/Accent 6
        //        break;
        //}

        //}

        /// <summary>
        /// Save to DOCX format
        /// </summary>
        /// <param name="path"></param>
        public void saveWordDoc(string path, bool showWord = false, string title = "")
        {
            System.Threading.Thread.Sleep(1000);
            try
            {
                string filename = path.Substring(path.LastIndexOf("\\") + 1);
                path = path.Substring(0, path.LastIndexOf("\\") + 1);
                filename = filename.Replace("<", " ")
                    .Replace(">", " ")
                    .Replace(":", " ")
                    .Replace("\"", " ")
                    .Replace("/", " ")
                    .Replace("|", " ")
                    .Replace("?", " ")
                    .Replace("*", " ");
                path += filename;

                if (!title.Equals("")) wordDoc.BuiltInDocumentProperties["Title"] = title; wordDoc.TablesOfContents[1].Update();
                wordDoc.SaveAs(path);
                if (showWord) wordDoc.Close(); wordApp.Visible = true; wordDoc = wordApp.Documents.Open(path); wordApp.WindowState = Word.WdWindowState.wdWindowStateMinimize; wordApp.WindowState = Word.WdWindowState.wdWindowStateMaximize;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                MessageBox.Show(path + " May be currently in use. Please exit any instances of it and click OK. \n\n Additional Details: \n" + e.Message);
                saveWordDoc(path, showWord, title);
            }
        }

        /// <summary>
        /// Closes Doc and Exits (Used when not showing doc afterwards)
        /// </summary>
        public void exitWord()
        {
            try
            {
                wordDoc.Close(Missing.Value, Missing.Value, Missing.Value);
                wordDoc = null;
                wordApp.Quit(Missing.Value, Missing.Value, Missing.Value);
                wordApp = null;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }


        public void debugText(string text)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            using (System.IO.StreamWriter outputFile = new System.IO.StreamWriter(desktopPath + @"\OfficeWrapperDebug.txt", true))
            {
                outputFile.WriteLine(text);
            }
        }

    }
}