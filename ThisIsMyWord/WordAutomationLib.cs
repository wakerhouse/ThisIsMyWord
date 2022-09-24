using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Documents;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;
using Word = Microsoft.Office.Interop.Word;

namespace ThisIsMyWord
{
    public class WordAutomationLib
    {
        //https://learn.microsoft.com/de-de/previous-versions/office/troubleshoot/office-developer/automate-word-create-file-using-visual-c
        //https://learn.microsoft.com/de-de/visualstudio/vsto/word-object-model-overview?source=recommendations&view=vs-2022&tabs=csharp

        //Start Word and create a new document.
        private Word._Application oWord;
        private Word._Document oDoc;
        private object oMissing = System.Reflection.Missing.Value;
        private object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

        public WordAutomationLib()
        {

        }

        public Word._Document OpenTemplate(string template)
        {
            System.Diagnostics.Process.Start("WINWORD");
            System.Diagnostics.Process.GetProcesses().Where(x => x.ProcessName == "WINWORD").First().Kill();

            oWord = new Word.Application();
            oWord.Visible = true;
            object oTemplate = template;
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
            ref oMissing, ref oMissing);

            return oDoc;
        }

        public Word._Document NewDocument()
        {

            oWord = new Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);

            return oDoc;
        }

        public bool SaveDocument(string filename = @"C:\temp\WordFile.docx"
            , Word.WdSaveFormat wdFormat = WdSaveFormat.wdFormatDocumentDefault)
        {

            //Saving
            oDoc.SaveAs2(filename, wdFormat);

            //Check if Document exist
            if (File.Exists(filename) == false)
                return false;

            return true;
        }

        public bool CloseDocument(bool saveChanges = true)
        {
            //Close document and save changes
            //oDoc.Close(saveChanges);
            oWord.Quit(saveChanges);

            return true;
        }

        public Word.Paragraph AddHeaderParagraph(string heading)
        {
            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara;
            oPara = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara.Range.Text = heading;
            oPara.Range.Font.Bold = 1;
            oPara.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara.Range.InsertParagraphAfter();

            return oPara;
        }

        public Word.Paragraph AddParagraph(string paragraph
            , float size = 11, int bold = 0, int italic = 0
            , Word.WdColor foreColor = WdColor.wdColorBlack
            , Word.WdUnderline underline = WdUnderline.wdUnderlineNone 
            , int spaceAfter = 0)
        {

            //Insert another paragraph.
            Word.Paragraph oPara;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.Text = paragraph;
            oPara.Range.Font.Color = foreColor;
            oPara.Range.Font.Size = size;
            oPara.Range.Font.Underline = underline;
            oPara.Range.Font.Bold = bold;
            oPara.Range.Font.Italic = italic;
            oPara.Format.SpaceAfter = spaceAfter;
            oPara.Range.InsertParagraphAfter();

            return oPara;
        }

        public Word.Table AddTable(string csvTableWithHeader, char separator = ';', int spaceAfter = 0)
        {
            //Insert a table, fill it with data, and make the first row
            //bold and italic.
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

            //Parse csv
            string[] lines = csvTableWithHeader.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            int rows = lines.Length;
            int cols = 0;
            //If lines count > 0
            if (lines.Length == 0)
                return null;
            //Check columns
            cols = lines[0].Split(separator).Length;
            oTable = oDoc.Tables.Add(wrdRng, rows, cols, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = spaceAfter;
            oTable.AllowPageBreaks = false;
            oTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth025pt;
            oTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oTable.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth025pt;


            for (int r = 0; r < rows; r++)
            {
                string[] content = lines[r].Split(separator);
                for (int c = 0; c < cols; c++)
                {
                    oTable.Cell(r+1, c+1).Range.Text = content[c];
                }
            }

            return oTable;
        }

        public void TableHeaderFormat(Word.Table table, int bold = 1, int italic = 0, Word.WdColor bgcolor = WdColor.wdColorGray30)
        {
            if (table.Rows.Count == 0)
                return;

            table.Rows[1].Range.Font.Bold = bold;
            table.Rows[1].Range.Font.Italic = italic;
            table.Rows[1].Range.Shading.BackgroundPatternColor = bgcolor;
        }

        public void TableColumnsWidth(Word.Table table, int column, int width)
        {
            table.Columns[column].Width = oWord.InchesToPoints(width); //Change width of columns
        }

        public void TableLineStyle(Word.Table table
            , Word.WdLineStyle style = Word.WdLineStyle.wdLineStyleSingle
            , Word.WdLineWidth width = Word.WdLineWidth.wdLineWidth025pt
            ,Word.WdColor color = Word.WdColor.wdColorBlack)
        {
            table.Borders.OutsideColor = color;
            table.Borders.InsideColor = color;
            table.Borders.OutsideLineStyle = style;
            table.Borders.OutsideLineWidth = width;
            table.Borders.InsideLineStyle = style;
            table.Borders.InsideLineWidth = width;
        }

        public void AddMultipleParagraphs(params string[] paragraphs)
        {
            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            double dPos = oWord.InchesToPoints(7);
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            foreach(var parapgraph in paragraphs)
            {
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.ParagraphFormat.SpaceAfter = 0;
                wrdRng.InsertAfter(parapgraph);
                wrdRng.InsertParagraphAfter();
            }                           
        }

        public void Collapse(Word.WdCollapseDirection direction)
        {
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            object oCollapse = direction;
            wrdRng.Collapse(ref oCollapse);
        }

        public void InsertPageBreak()
        {
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Word.WdBreakType.wdPageBreak;
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertBreak(ref oPageBreak);
        }

        public Word.InlineShape AddPicture(string filename)
        {
            Word.InlineShape image;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            image = oDoc.InlineShapes.AddPicture(filename,false,true,wrdRng);

            return image;
        }

        public void AddChart(double[] data, string title, string xLabel, string yLabel, float width, float height)
        {
            //https://stackoverflow.com/questions/3684103/how-to-add-office-graph-in-word
            //Insert a chart.
            Word.InlineShape oShape;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oShape = oDoc.InlineShapes.AddChart(Microsoft.Office.Core.XlChartType.xlXYScatterLines,wrdRng);
            //Size
            oShape.Width = width;
            oShape.Height = height;
            //Access to the data
            oShape.Chart.ChartData.Activate();
            Word.Chart objChart = oShape.Chart;
            Workbook book = objChart.ChartData.Workbook;
            Worksheet sheet = book.Worksheets["Tabelle1"];
            //Title
            objChart.ChartTitle.Text = title;
            //Label
            sheet.Cells[1, 1] = xLabel;
            sheet.Cells[1, 2] = yLabel;
            //DaTa
            for (int i = 0; i < data.Length; i++)
            {
                sheet.Cells[i + 2, 1] = i;
                sheet.Cells[i + 2, 2] = data[i];
            }

            oShape.Chart.SetSourceData("'Tabelle1'!A1:B" + data.Length+1);

            //book.Save();
            book.Close();

        }

    }

}
