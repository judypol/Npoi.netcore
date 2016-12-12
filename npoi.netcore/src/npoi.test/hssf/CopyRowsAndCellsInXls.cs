using System.IO;
using npoi.core.Corefx;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace npoi.test.hssf
{
    public class CopyRowsAndCellsInXls
    {
        [Fact]
        public void Main()
        {
            InitializeWorkbook();

            ISheet s = hssfworkbook.GetSheetAt(0);

            ICell cell = s.GetRow(4).GetCell(1);
            cell.CopyCellTo(3); //copy B5 to D5

            IRow c = s.GetRow(3);
            c.CopyCell(0, 1);   //copy A4 to B4

            s.CopyRow(0, 1);     //copy row A to row B, original row B will be moved to row C automatically
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
            using (var fs = File.OpenRead(@"Data\test.xls"))
            {

                hssfworkbook = new HSSFWorkbook(fs);
            }
        }
    }
}
