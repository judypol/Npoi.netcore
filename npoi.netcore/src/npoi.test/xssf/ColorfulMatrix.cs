﻿using System.IO;
using npoi.core.Corefx;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace npoi.test.xssf
{
    public class ColorfulMatrix
    {
        [Fact]
        public void Main()
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            int x = 1;
            for (int i = 0; i < 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 15; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (x % 2 == 0)
                    {
                        //fill background with blue
                        ICellStyle style1 = workbook.CreateCellStyle();
                        style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Blue.Index2;
                        style1.FillPattern = FillPattern.SolidForeground;
                        cell.CellStyle = style1;
                    }
                    else
                    {
                        //fill background with yellow
                        ICellStyle style1 = workbook.CreateCellStyle();
                        style1.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index2;
                        style1.FillPattern = FillPattern.SolidForeground;
                        cell.CellStyle = style1;
                    }
                    x++;
                }
            }
            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
