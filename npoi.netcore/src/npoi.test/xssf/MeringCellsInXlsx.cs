using System.IO;
using npoi.core.Corefx;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace npoi.test.xssf
{
    public class MeringCellsInXlsx
    {
        [Fact]
        public void Main()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("PictureSheet");

            IRow row = sheet.CreateRow(1);
            ICell cell = row.CreateCell(1);
            cell.SetCellValue(new XSSFRichTextString("This is a test of merging"));

            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 1, 2));

            FileStream sw = File.OpenWrite("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
