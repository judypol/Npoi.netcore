using System.IO;
using npoi.core.Corefx;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace npoi.test.xssf
{
    public class ProtectSheetInXlsx
    {
        [Fact]
        public void Main()
        {
            IWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)workbook.CreateSheet("Sheet A1");

            sheet1.LockFormatRows();
            sheet1.LockFormatCells();
            sheet1.LockFormatColumns();
            sheet1.LockDeleteColumns();
            sheet1.LockDeleteRows();
            sheet1.LockInsertHyperlinks();
            sheet1.LockInsertColumns();
            sheet1.LockInsertRows();
            sheet1.ProtectSheet("password");


            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
