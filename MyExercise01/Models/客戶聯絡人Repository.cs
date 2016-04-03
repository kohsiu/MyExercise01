using System;
using System.Linq;
using System.Collections.Generic;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.IO;

namespace MyExercise01.Models
{   
	public  class 客戶聯絡人Repository : EFRepository<客戶聯絡人>, I客戶聯絡人Repository
	{
        public override IQueryable<客戶聯絡人> All()
        {
            return base.All().Include("客戶聯絡人.客戶資料").Where(w => w.是否已刪除 == false).OrderBy(o => o.客戶Id);
        }

        public override void Delete(客戶聯絡人 entity)
        {
            entity.是否已刪除 = true;
        }

        public 客戶聯絡人 Find(int id)
        {
            return this.All().Include("客戶聯絡人.客戶資料").FirstOrDefault(w => w.Id == id);
        }

        public IQueryable<客戶聯絡人> Search (string searchName, string searchCusName)
        {
            return this.All().Where(w => w.姓名.Contains(searchName) &&
                                    w.客戶資料.客戶名稱.Contains(searchCusName));
        }

        public byte[] ExprotExcel(List<客戶聯絡人> data)
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
            headerRow.CreateCell(0).SetCellValue("職稱");
            headerRow.Cells[0].CellStyle = styletital;
            headerRow.CreateCell(1).SetCellValue("姓名");
            headerRow.Cells[1].CellStyle = styletital;
            headerRow.CreateCell(2).SetCellValue("Email");
            headerRow.Cells[2].CellStyle = styletital;
            headerRow.CreateCell(3).SetCellValue("手機");
            headerRow.Cells[3].CellStyle = styletital;
            headerRow.CreateCell(4).SetCellValue("電話");
            headerRow.Cells[4].CellStyle = styletital;
            headerRow.CreateCell(5).SetCellValue("客戶名稱");
            headerRow.Cells[5].CellStyle = styletital;
            #endregion

            #region DetailRow
            int rowIndex = 1;
            foreach (var item in data)
            {
                XSSFRow dataRow = (XSSFRow)sheet.CreateRow(rowIndex);

                dataRow.CreateCell(0).SetCellValue(item.職稱);
                dataRow.Cells[0].CellStyle = stylecon;
                dataRow.CreateCell(1).SetCellValue(item.姓名);
                dataRow.Cells[1].CellStyle = stylecon;
                dataRow.CreateCell(2).SetCellValue(item.Email);
                dataRow.Cells[2].CellStyle = stylecon;
                dataRow.CreateCell(3).SetCellValue(item.手機);
                dataRow.Cells[3].CellStyle = stylecon;
                dataRow.CreateCell(4).SetCellValue(item.電話);
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

    public  interface I客戶聯絡人Repository : IRepository<客戶聯絡人>
	{

	}
}