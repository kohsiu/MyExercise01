using System;
using System.Linq;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.IO;

namespace MyExercise01.Models
{   
	public  class View_客戶資訊清單Repository : EFRepository<View_客戶資訊清單>, IView_客戶資訊清單Repository
	{
        public byte[] ExprotExcel(IList<View_客戶資訊清單> data)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet("Export");

            XSSFCellStyle styletital = (XSSFCellStyle)workbook.CreateCellStyle();
            styletital.Alignment = HorizontalAlignment.Center;
            styletital.BorderBottom = BorderStyle.Thin;
            styletital.BorderLeft = BorderStyle.Thin;
            styletital.BorderRight = BorderStyle.Thin;
            styletital.BorderTop = BorderStyle.Thin;
            styletital.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            styletital.FillPattern = FillPattern.SolidForeground;

            XSSFCellStyle stylecon = (XSSFCellStyle)workbook.CreateCellStyle();
            stylecon.Alignment = HorizontalAlignment.Left; //對齊
            stylecon.BorderBottom = BorderStyle.Thin;
            stylecon.BorderLeft = BorderStyle.Thin;
            stylecon.BorderRight = BorderStyle.Thin;
            stylecon.BorderTop = BorderStyle.Thin;

            #region HeaderRow
            XSSFRow headerRow = (XSSFRow)sheet.CreateRow(0);

            headerRow.CreateCell(0).SetCellValue("客戶名稱");
            headerRow.Cells[0].CellStyle = styletital;
            headerRow.CreateCell(1).SetCellValue("客戶銀行資訊數量");
            headerRow.Cells[1].CellStyle = styletital;
            headerRow.CreateCell(2).SetCellValue("客戶聯絡人數量");
            headerRow.Cells[2].CellStyle = styletital;
            #endregion

            #region DetailRow
            int rowIndex = 1;
            foreach (var item in data)
            {
                XSSFRow dataRow = (XSSFRow)sheet.CreateRow(rowIndex);

                dataRow.CreateCell(0).SetCellValue(item.客戶名稱);
                dataRow.Cells[0].CellStyle = stylecon;
                dataRow.CreateCell(1).SetCellValue(item.客戶銀行資訊數量.ToString());
                dataRow.Cells[1].CellStyle = stylecon;
                dataRow.CreateCell(2).SetCellValue(item.客戶聯絡人數量.ToString());
                dataRow.Cells[2].CellStyle = stylecon;

                rowIndex++;
            }
            #endregion

            for (int i = 0; i < sheet.GetRow(0).Cells.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            // code to create workbook 

            byte[] fileContents = null;
            using (var memoryStream = new MemoryStream())
            {
                workbook.Write(memoryStream);
                fileContents = memoryStream.ToArray();

            }
            return fileContents;
        }
    }

	public  interface IView_客戶資訊清單Repository : IRepository<View_客戶資訊清單>
	{

	}
}