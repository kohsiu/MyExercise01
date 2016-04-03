using System;
using System.Linq;
using System.Collections.Generic;
using System.Web.Mvc;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;

namespace MyExercise01.Models
{
    public class 客戶資料Repository : EFRepository<客戶資料>, I客戶資料Repository
    {
        public override IQueryable<客戶資料> All()
        {
            return base.All().Where(w => w.是否已刪除 == false).OrderBy(o => o.Id);
        }

        public override void Delete(客戶資料 entity)
        {
            entity.是否已刪除 = true;
        }

        public 客戶資料 Find(int id)
        {
            return this.All().Include("客戶資料.客戶聯絡人").FirstOrDefault(w => w.Id == id);
        }

        public IQueryable<客戶資料> SearchCusName(string keyword, string kind)
        {
            return this.All().Where(w => w.客戶名稱.Contains(keyword) && w.客戶分類 == kind);
        }

        public List<客戶分類Models> CreateCusKindData()
        {
            List<客戶分類Models> dropdownData = new List<客戶分類Models>();

            for (int i = 0; i < 10; i++)
            {
                string stri = (i + 1).ToString();
                dropdownData.Add(new 客戶分類Models()
                {
                    Kind = "分類" + stri
                });
            }

            return dropdownData;
        }

        public byte[] ExprotExcel(IList<客戶資料> data)
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
            headerRow.CreateCell(1).SetCellValue("統一編號");
            headerRow.Cells[1].CellStyle = styletital;
            headerRow.CreateCell(2).SetCellValue("電話");
            headerRow.Cells[2].CellStyle = styletital;
            headerRow.CreateCell(3).SetCellValue("傳真");
            headerRow.Cells[3].CellStyle = styletital;
            headerRow.CreateCell(4).SetCellValue("地址");
            headerRow.Cells[4].CellStyle = styletital;
            headerRow.CreateCell(5).SetCellValue("Email");
            headerRow.Cells[5].CellStyle = styletital;
            headerRow.CreateCell(6).SetCellValue("客戶分類");
            headerRow.Cells[6].CellStyle = styletital;
            #endregion

            #region DetailRow
            int rowIndex = 1;
            foreach (var item in data)
            {
                XSSFRow dataRow = (XSSFRow)sheet.CreateRow(rowIndex);

                dataRow.CreateCell(0).SetCellValue(item.客戶名稱);
                dataRow.Cells[0].CellStyle = stylecon;
                dataRow.CreateCell(1).SetCellValue(item.統一編號);
                dataRow.Cells[1].CellStyle = stylecon;
                dataRow.CreateCell(2).SetCellValue(item.電話);
                dataRow.Cells[2].CellStyle = stylecon;
                dataRow.CreateCell(3).SetCellValue(item.傳真);
                dataRow.Cells[3].CellStyle = stylecon;
                dataRow.CreateCell(4).SetCellValue(item.地址);
                dataRow.Cells[4].CellStyle = stylecon;
                dataRow.CreateCell(5).SetCellValue(item.Email);
                dataRow.Cells[5].CellStyle = stylecon;
                dataRow.CreateCell(6).SetCellValue(item.客戶分類);
                dataRow.Cells[6].CellStyle = stylecon;

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

    public interface I客戶資料Repository : IRepository<客戶資料>
    {

    }
}