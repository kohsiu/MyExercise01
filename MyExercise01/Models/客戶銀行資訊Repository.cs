using System;
using System.Linq;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.IO;

namespace MyExercise01.Models
{   
	public  class 客戶銀行資訊Repository : EFRepository<客戶銀行資訊>, I客戶銀行資訊Repository
	{
        public override IQueryable<客戶銀行資訊> All()
        {

            return base.All().Where(w => w.是否已刪除 == false).OrderBy(o => o.客戶Id);
            
        }

        public override void Delete(客戶銀行資訊 entity)
        {
            base.Delete(entity);    
        }

        public 客戶銀行資訊 Find(int id)
        {
            return this.All().FirstOrDefault(w => w.Id == id);
        }

        public IQueryable<客戶銀行資訊> SearchCusName(string keyword)
        {
            return this.All().Include("客戶銀行資訊.客戶資料").Where(w => w.帳戶名稱.Contains(keyword));
        }

        public byte[] ExprotExcel(List<客戶銀行資訊> data)
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
            headerRow.CreateCell(0).SetCellValue("銀行名稱");
            headerRow.Cells[0].CellStyle = styletital;
            headerRow.CreateCell(1).SetCellValue("銀行代碼");
            headerRow.Cells[1].CellStyle = styletital;
            headerRow.CreateCell(2).SetCellValue("分行代碼");
            headerRow.Cells[2].CellStyle = styletital;
            headerRow.CreateCell(3).SetCellValue("帳戶名稱");
            headerRow.Cells[3].CellStyle = styletital;
            headerRow.CreateCell(4).SetCellValue("帳戶號碼");
            headerRow.Cells[4].CellStyle = styletital;
            headerRow.CreateCell(5).SetCellValue("客戶名稱");
            headerRow.Cells[5].CellStyle = styletital;
            #endregion

            #region DetailRow
            int rowIndex = 1;
            foreach (var item in data)
            {
                XSSFRow dataRow = (XSSFRow)sheet.CreateRow(rowIndex);

                dataRow.CreateCell(0).SetCellValue(item.銀行名稱);
                dataRow.Cells[0].CellStyle = stylecon;
                dataRow.CreateCell(1).SetCellValue(item.銀行代碼);
                dataRow.Cells[1].CellStyle = stylecon;
                dataRow.CreateCell(2).SetCellValue(item.分行代碼.ToString());
                dataRow.Cells[2].CellStyle = stylecon;
                dataRow.CreateCell(3).SetCellValue(item.帳戶名稱);
                dataRow.Cells[3].CellStyle = stylecon;
                dataRow.CreateCell(4).SetCellValue(item.帳戶號碼);
                dataRow.Cells[4].CellStyle = stylecon;
                dataRow.CreateCell(5).SetCellValue(item.客戶資料.客戶名稱);
                dataRow.Cells[5].CellStyle = stylecon;

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

	public  interface I客戶銀行資訊Repository : IRepository<客戶銀行資訊>
	{

	}
}