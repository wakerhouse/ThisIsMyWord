using System;
using System.Reflection;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace ThisIsMyWord
{
    public class WordAutomationLib
    {
        //https://learn.microsoft.com/de-de/previous-versions/office/troubleshoot/office-developer/automate-word-create-file-using-visual-c

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

        public void AddStartParagraph(string heading)
        {
            //Insert a paragraph at the beginning of the document.
            Word.Paragraph oPara;
            oPara = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara.Range.Text = heading;
            oPara.Range.Font.Bold = 1;
            oPara.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara.Range.InsertParagraphAfter();

        }

        public void AddEndParagraph(string heading)
        {

            //Insert a paragraph at the end of the document.
            Word.Paragraph oPara;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.Text = heading;
            oPara.Format.SpaceAfter = 6;
            oPara.Range.InsertParagraphAfter();

        }


        public void AddParagraph(string paragraph)
        {

            //Insert another paragraph.
            Word.Paragraph oPara;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara.Range.Text = paragraph;
            oPara.Range.Font.Bold = 0;
            oPara.Format.SpaceAfter = 0;
            oPara.Range.InsertParagraphAfter();

        }

        public void AddTable(string csvTableWithHeader, char separator = ';')
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
                return;
            //Check columns
            cols = lines[0].Split(separator).Length;
            oTable = oDoc.Tables.Add(wrdRng, rows, cols, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 0;


            for (int r = 0; r < rows; r++)
            {
                string[] content = lines[r].Split(separator);
                for (int c = 0; c < cols; c++)
                {
                    oTable.Cell(r+1, c+1).Range.Text = content[c];
                }
            }


            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;
            //oTable.Columns[1].Width = oWord.InchesToPoints(2); //Change width of columns 1 & 2
            //oTable.Columns[2].Width = oWord.InchesToPoints(3);
        }


        public void Add()
        {
            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            object oPos;
            double dPos = oWord.InchesToPoints(7);
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            do
            {
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.ParagraphFormat.SpaceAfter = 0;
                wrdRng.InsertAfter("A line of text");
                wrdRng.InsertParagraphAfter();
                oPos = wrdRng.get_Information
                                       (Word.WdInformation.wdVerticalPositionRelativeToPage);
            }
            while (dPos >= Convert.ToDouble(oPos));
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Word.WdBreakType.wdPageBreak;
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertBreak(ref oPageBreak);
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            wrdRng.InsertParagraphAfter();
        }

        public void InsertPageBreak()
        {
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            object oCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = Word.WdBreakType.wdPageBreak;
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertBreak(ref oPageBreak);
        }

        public void AddChart()
        {
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //Insert a chart.
            Word.InlineShape oShape;
            object oClassType = "MSGraph.Chart.8";
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing);

            //Demonstrate use of late bound oChart and oChartApp objects to
            //manipulate the chart object with MSGraph.
            object oChart;
            object oChartApp;
            oChart = oShape.OLEFormat.Object;
            oChartApp = oChart.GetType().InvokeMember("Application",
            BindingFlags.GetProperty, null, oChart, null);

            //Change the chart type to Line.
            object[] Parameters = new Object[1];
            Parameters[0] = 4; //xlLine = 4
            oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
            null, oChart, Parameters);

            //Update the chart image and quit MSGraph.
            oChartApp.GetType().InvokeMember("Update",
            BindingFlags.InvokeMethod, null, oChartApp, null);
            oChartApp.GetType().InvokeMember("Quit",
            BindingFlags.InvokeMethod, null, oChartApp, null);
            //... If desired, you can proceed from here using the Microsoft Graph 
            //Object model on the oChart and oChartApp objects to make additional
            //changes to the chart.

            //Set the width of the chart.
            oShape.Width = oWord.InchesToPoints(6.25f);
            oShape.Height = oWord.InchesToPoints(3.57f);
        }


    }
}
