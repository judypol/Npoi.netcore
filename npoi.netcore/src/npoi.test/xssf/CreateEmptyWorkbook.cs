using System.IO;
using npoi.core.Corefx;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace npoi.test.xssf
{
    public class CreateEmptyWorkbook
    {
        [Fact]
        public void Main()
        {
            IWorkbook workbook = new XSSFWorkbook();
            workbook.CreateSheet("Sheet A1");
            workbook.CreateSheet("Sheet A2");
            workbook.CreateSheet("Sheet A3");

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
