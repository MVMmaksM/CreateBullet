using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;

namespace CreateBullet
{
    internal class Bulletines
    {
        public ExcelPackage? Create(DataTable list, DataTable listVigr)
        {
            try
            {
                Console.WriteLine($"\nНачинаю объединять!");

                var query = list.AsEnumerable().Join(listVigr.AsEnumerable(), l => l["kod"], v => v["kod"],
                   (l, v) => new
                   {
                       Poryadok = l["porydok"],
                       Name = l["name"],
                       Ufa = v["ufa"],
                       Ijevsk = v["ijevsk"],
                       Perm = v["perm"],
                       Orenburg = v["orenburg"],
                       Kurgan = v["kurgan"],
                       Ekat = v["ekaterinburg"],
                       Tumen = v["tumen"],
                       Hanty = v["hanty"],
                       Salehard = v["salehard"],
                       Chelyabinsk = v["chelyabinsk"]
                   }).OrderBy(a => a.Poryadok);


                ExcelPackage excelPackage = new();
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add("Лист 1");

                excelWorksheet.Column(1).Width = 50;
                excelWorksheet.Row(1).Height = 20;

                excelWorksheet.Cells[1, 1, 2, 1].Merge = true;
                excelWorksheet.Cells[1, 1, 2, 1].Value = "НАИМЕНОВАНИЕ";
                excelWorksheet.Cells[1, 1, 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 1, 2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                excelWorksheet.Cells[1, 2, 1, 5].Merge = true;
                excelWorksheet.Cells[1, 2, 1, 5].Value = "Приволжский федеральный округ";
                excelWorksheet.Cells[1, 2, 1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 2, 1, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                excelWorksheet.Cells[1, 6, 1, 11].Merge = true;
                excelWorksheet.Cells[1, 6, 1, 11].Value = "Уральский федеральный округ";
                excelWorksheet.Cells[1, 6, 1, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 6, 1, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                excelWorksheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelWorksheet.Cells[1, 1, 2, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelWorksheet.Cells[1, 1, 2, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelWorksheet.Cells[2, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                excelWorksheet.Cells[1, 2, 1, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                excelWorksheet.Cells[1, 2, 1, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                excelWorksheet.Cells[1, 2, 1, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelWorksheet.Cells[1, 2, 1, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                excelWorksheet.Cells[2, 2, 2, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                excelWorksheet.Cells[2, 2, 2, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                excelWorksheet.Cells[2, 2].Value = "Уфа";
                excelWorksheet.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 3].Value = "Ижевск";
                excelWorksheet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 4].Value = "Пермь";
                excelWorksheet.Cells[2, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 5].Value = "Оренбург";
                excelWorksheet.Cells[2, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 6].Value = "Курган";
                excelWorksheet.Cells[2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 7].Value = "Екатеринбург";
                excelWorksheet.Cells[2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 8].Value = "Тюмень";
                excelWorksheet.Cells[2, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 9].Value = "Ханты-Мансийск";
                excelWorksheet.Cells[2, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 10].Value = "Салехард";
                excelWorksheet.Cells[2, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[2, 11].Value = "Челябинск";
                excelWorksheet.Cells[2, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int rowCount = 3;

                foreach (var row in query)
                {
                    excelWorksheet.Cells[rowCount, 1].Value = row.Name;
                    excelWorksheet.Cells[rowCount, 2].Value = row.Ufa;
                    excelWorksheet.Cells[rowCount, 3].Value = row.Ijevsk;
                    excelWorksheet.Cells[rowCount, 4].Value = row.Perm;
                    excelWorksheet.Cells[rowCount, 5].Value = row.Orenburg;
                    excelWorksheet.Cells[rowCount, 6].Value = row.Kurgan;
                    excelWorksheet.Cells[rowCount, 7].Value = row.Ekat;
                    excelWorksheet.Cells[rowCount, 8].Value = row.Tumen;
                    excelWorksheet.Cells[rowCount, 9].Value = row.Hanty;
                    excelWorksheet.Cells[rowCount, 10].Value = row.Salehard;
                    excelWorksheet.Cells[rowCount, 11].Value = row.Chelyabinsk;
                    rowCount++;
                }

                return excelPackage;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка: {ex.Message + ex.StackTrace}");
                return null;
            }            
        }
    }
}
