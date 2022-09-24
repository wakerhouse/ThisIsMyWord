using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Reflection;

namespace WordTest
{
    internal class Program
    {
        static void Main(string[] args)
        {

            ThisIsMyWord.WordAutomationLib wordAutomationLib = new ThisIsMyWord.WordAutomationLib();
            FileInfo fi = new FileInfo(Assembly.GetExecutingAssembly().FullName);
            wordAutomationLib.OpenTemplate(fi.DirectoryName + @"\ISO_Vorlage.dotx");
            wordAutomationLib.AddHeaderParagraph("Hallo Test");
            wordAutomationLib.AddParagraph("Table 1");
            string table = "1;2;3;4;5\r\newq;qwe;qwe;qwe;qwe";
            var oTable = wordAutomationLib.AddTable(table);
            wordAutomationLib.TableHeaderFormat(oTable, 0, 3, WdColor.wdColorGray20);
            wordAutomationLib.AddPicture(fi.DirectoryName + @"\image.jpg");
            wordAutomationLib.AddMultipleParagraphs("Was geht?", "Haha na klar");
            wordAutomationLib.InsertPageBreak();
            wordAutomationLib.AddHeaderParagraph("Hallo Test");
            wordAutomationLib.AddChart(new double[] {3,7,6,3,6,8,2,5}, "Data", "inc", "rnd", 400,200);
            wordAutomationLib.SaveDocument(@"C:\temp\testZert.docx", WdSaveFormat.wdFormatDocumentDefault);
            wordAutomationLib.SaveDocument(@"C:\temp\testZert.pdf", WdSaveFormat.wdFormatPDF);
            wordAutomationLib.CloseDocument(false);
        }
    }
}
