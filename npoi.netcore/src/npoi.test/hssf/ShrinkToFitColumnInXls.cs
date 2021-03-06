﻿using System.IO;
using npoi.core.Corefx;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace npoi.test.hssf
{
    public class ShrinkToFitColumnInXls
    {
        [Fact]
        public void Main()
        {
            InitializeWorkbook();

            ISheet sheet = hssfworkbook.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            //create cell value
            ICell cell1 = row.CreateCell(0);
            cell1.SetCellValue("This is a test");
            //apply ShrinkToFit to cellstyle
            ICellStyle cellstyle1 = hssfworkbook.CreateCellStyle();
            cellstyle1.ShrinkToFit = true;
            cell1.CellStyle = cellstyle1;
            //create cell value
            row.CreateCell(1).SetCellValue("Hello World");
            row.GetCell(1).CellStyle = cellstyle1;

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
