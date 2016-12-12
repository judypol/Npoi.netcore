using System.IO;
using npoi.core.Corefx;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace npoi.test.xssf
{
    public class BigGridTest
    {
        [Fact]
        public void Main()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("Sheet1");

            for (int rownum = 0; rownum < 10000; rownum++)
            {
                IRow row = worksheet.CreateRow(rownum);
                for (int celnum = 0; celnum < 20; celnum++)
                {
                    ICell Cell = row.CreateCell(celnum);
                    Cell.SetCellValue("Cell: Row-" + rownum + ";CellNo:" + celnum);
                }
            }

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
