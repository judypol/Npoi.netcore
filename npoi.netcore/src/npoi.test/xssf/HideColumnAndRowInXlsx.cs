using System.IO;
using npoi.core.Corefx;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace npoi.test.xssf
{
    public class HideColumnAndRowInXlsx
    {
        public void Main()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s = workbook.CreateSheet("Sheet1");
            IRow r1 = s.CreateRow(0);
            IRow r2 = s.CreateRow(1);
            IRow r3 = s.CreateRow(2);
            IRow r4 = s.CreateRow(3);
            IRow r5 = s.CreateRow(4);

            //hide IRow 2
            r2.ZeroHeight = true;

            //hide column C
            s.SetColumnHidden(2, true);
            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
