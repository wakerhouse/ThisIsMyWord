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
            wordAutomationLib.AddStartParagraph("Hallo Test");
            wordAutomationLib.AddEndParagraph("Ende");
            wordAutomationLib.AddParagraph("Table 1");
            string table = "1;2;3;4;5\r\newq;qwe;qwe;qwe;qwe";
            wordAutomationLib.AddTable(table);
            wordAutomationLib.Add();
            wordAutomationLib.InsertPageBreak();
            wordAutomationLib.AddStartParagraph("Hallo Test");

        }
    }
}
