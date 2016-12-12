using System;
using System.IO;
using npoi.core.Corefx;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace npoi.test.xssf
{
    public class CreateCustomProperties
    {
        [Fact]
        public void Main()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            POIXMLProperties props = workbook.GetProperties();
            props.CoreProperties.Creator = "NPOI 2.0.5";
            props.CoreProperties.Created = DateTime.Now;
            if (!props.CustomProperties.Contains("NPOI Team"))
                props.CustomProperties.AddProperty("NPOI Team", "Hello World!");

            FileStream sw = File.Create("test.xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
