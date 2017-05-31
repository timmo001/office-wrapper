using System;

namespace Word
{
    class Program
    {
        static void Main(string[] args)
        {
            //OfficeWrapper.Document wordDoc = new OfficeWrapper.Document();

            //wordDoc.createWordDoc(@"\\bd-gsc-san2\iseries_pdf$\PIDTemplate\PIDTemplate.dotx");

            //wordDoc.setFooter("D3213-2 Title");

            //wordDoc.addToTable(1, 2, 1, "Timmo");
            //wordDoc.addToTable(1, 4, 1, "Timmo");
            //wordDoc.addToTable(1, 6, 1, "Timmo");
            //wordDoc.addToTable(1, 8, 1, "Timmo");
            //wordDoc.addToTable(1, 10, 1, "Timmo");
            //wordDoc.addToTable(1, 12, 1, "Timmo");
            //wordDoc.addToTable(1, 14, 1, "Timmo");

            //wordDoc.addParagraph("Test\nTest 2\nTest 3", "two");

            //string tableNo = wordDoc.addTable(1, 1);
            //Console.WriteLine(tableNo);
            //wordDoc.addToTable(int.Parse(tableNo), 1, 1, "Test");

            //wordDoc.saveWordDoc(@"C:\Reports\" + DateTime.UtcNow + "TestRun.docx", true);

            OfficeWrapper.PowerPoint ppt = new OfficeWrapper.PowerPoint();

            ppt.createPPT();

            ppt.addBlankSlide(1);

            ppt.addTable(8, 8, "670", "300", "120", "25");

            ppt.styleTable(1);

            ppt.addTableContent(1, 1, "test");

            ppt.addShape((ppt.getCellX(1, 2) + (ppt.getCellWidth(1, 2) / 2) - 5).ToString(), (ppt.getCellY(1, 2) + (ppt.getCellHeight(1, 2) / 2) - 5).ToString(), "10", "10");

            ppt.savePPT(@"C:\Reports\test.pptx", true);

        }
    }
}
