using System.IO;
using npoi.core.Corefx;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace npoi.test.xssf
{
    public class CreateHeaderFooterInXlsx
    {
        [Fact]
        public void Main()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s1 = workbook.CreateSheet("Sheet1");
            s1.CreateRow(0).CreateCell(1).SetCellValue(123);

            //set header text
            s1.Header.Left = HSSFHeader.Page;   //Page is a static property of HSSFHeader and HSSFFooter
            s1.Header.Center = "This is a test sheet";
            //set footer text
            s1.Footer.Left = "Copyright NPOI Team";
            s1.Footer.Right = "created by Tony Qu（瞿杰）";
            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
