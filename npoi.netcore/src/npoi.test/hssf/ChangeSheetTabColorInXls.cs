using System.IO;
using npoi.core.Corefx;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using Xunit;

namespace npoi.test.hssf
{
    public class ChangeSheetTabColorInXls
    {
        [Fact]
        public void Main()
        {
            InitializeWorkbook();

            ISheet sheet1 = hssfworkbook.CreateSheet("Sheet1");
            sheet1.TabColorIndex = HSSFColor.Red.Index;
            ISheet sheet2 = hssfworkbook.CreateSheet("Sheet2");
            sheet2.TabColorIndex = HSSFColor.Blue.Index;
            ISheet sheet3 = hssfworkbook.CreateSheet("Sheet3");
            sheet3.TabColorIndex = HSSFColor.Aqua.Index;

            WriteToFile();
        }


        static HSSFWorkbook hssfworkbook;

        static void WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xls", FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }

        static void InitializeWorkbook()
        {
            hssfworkbook = new HSSFWorkbook();

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
