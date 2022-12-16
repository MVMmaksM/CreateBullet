using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BulletCreate.Model;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BulletCreate.Services
{
    internal class FileServices
    {
        private static bool IsValidNumberColumn(IEnumerable<ModelNameColumnDataMarts> modelNameColumnDataMarts)
        {
            foreach (var dataMarts in modelNameColumnDataMarts)
            {
                if (dataMarts.NumberColumn is 0)
                {
                    Console.WriteLine($"В витринах данных отсутвует столбец с названием: {dataMarts.NameColumn}");
                    return false;
                }
            }

            return true;
        }
        private static IEnumerable<ModelNameColumnDataMarts> GetNumberColumn(IEnumerable<ModelNameColumnDataMarts> nameNumberColumn, ExcelWorksheet sheet)
        {

            if (sheet.Cells[1, 1].Value.ToString() == "Наименование" && sheet.Cells[1, 2].Value.ToString() == "Код")
            {
                sheet.Cells[2, 1].Value = sheet.Cells[1, 1].Value.ToString();
                sheet.Cells[2, 2].Value = sheet.Cells[1, 2].Value.ToString();
            }


            for (byte i = 1; i < sheet.Dimension.End.Column; i++)
            {
                foreach (var dataMarts in nameNumberColumn)
                {
                    if (sheet.Cells[2, i].Value.ToString() == dataMarts.NameColumn)
                    {
                        dataMarts.NumberColumn = i;
                        break;
                    }
                }
            }

            return nameNumberColumn;
        }
        private static string? GetFileNameNomenclature(string pathFile)
        {
            try
            {
                string[] files = Directory.GetFiles(pathFile);

                if (files.Length > 1)
                {
                    Console.WriteLine("\n****************Ошибка!**************");
                    Console.WriteLine("В директории должен быть только один файл с порядком позиций!");
                    return null;
                }
                else if (files.Length == 0)
                {
                    Console.WriteLine("\n****************Ошибка!**************");
                    Console.WriteLine("Отсутствует файл с порядком позиций!");
                    return null;
                }

                return files[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n****************Ошибка!**************");
                Console.WriteLine($"{ex.Message + ex.StackTrace}");
                return null;
            }
        }
        private static string? GetFileNameDataMarts(string pathFile)
        {
            try
            {
                string[] files = Directory.GetFiles(pathFile);

                if (files.Length > 1)
                {
                    Console.WriteLine("\n****************Ошибка!**************");
                    Console.WriteLine("В директории должен быть только один файл с выгрузкой!");
                    return null;
                }
                else if (files.Length == 0)
                {
                    Console.WriteLine("\n****************Ошибка!**************");
                    Console.WriteLine("Отсутствует файл с выгрузкой из витрин данных!");
                    return null;
                }

                return files[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n****************Ошибка!**************");
                Console.WriteLine($"{ex.Message + ex.StackTrace}");
                return null;
            }
        }
        public static IEnumerable<ModelDataMarts>? LoadInDataMarts(string pathDirectory)
        {
            string? pathFile = GetFileNameDataMarts(pathDirectory);

            if (pathFile == null)
            {
                return null;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(pathFile);
            ExcelWorksheet sheet = package.Workbook.Worksheets[0];

            IEnumerable<ModelNameColumnDataMarts> nameColumnDataMarts = GetNumberColumn(ModelNameColumnDataMarts.CreateNameColumnDataMarts(), sheet);

            if (!IsValidNumberColumn(nameColumnDataMarts))
            {
                return null;
            }

            List<ModelDataMarts> dataMarts = new List<ModelDataMarts>();
            int k = 0;

            try
            {

                for (int i = 119; i <sheet.Dimension.End.Row; i++)
                {
                    var row = sheet.Cells[i, 1, i, sheet.Dimension.End.Column];

                    ModelDataMarts rowDatatMarts = new();

                    rowDatatMarts.Name = row[i, 1].Value.ToString();
                    rowDatatMarts.Kod = int.Parse(row[i, 2].Value.ToString());
                    rowDatatMarts.Ufa = row[i, 3].Value?.ToString();
                    rowDatatMarts.Ijevsk = row[i, 4].Value?.ToString();
                    rowDatatMarts.Perm = row[i, 5].Value?.ToString();
                    rowDatatMarts.Orenburg = row[i, 6].Value?.ToString();
                    rowDatatMarts.Kurgan = row[i, 7].Value?.ToString();
                    rowDatatMarts.Ekaterinburg = row[i, 8].Value?.ToString();
                    rowDatatMarts.Tumen = row[i, 9].Value?.ToString();
                    rowDatatMarts.Hanty = row[i, 10].Value?.ToString();
                    rowDatatMarts.Salehard = row[i, 11].Value?.ToString();
                    rowDatatMarts.Chelyabinsk = row[i, 12].Value?.ToString();

                    dataMarts.Add(rowDatatMarts);
                    k++;
                }

                return dataMarts;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + $"итерация{k} ");
                return null;
            }

        }
        public static IEnumerable<ModelNomenclature>? LoadInNomenclature(string pathDirectory)
        {
            string? pathFile = GetFileNameNomenclature(pathDirectory);

            if (pathFile == null)
            {
                return null;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(pathFile);
            ExcelWorksheet sheet = package.Workbook.Worksheets[0];

            List<ModelNomenclature> modelNomenclatures = new();

            for (int i = 1; i <= sheet.Dimension.End.Row; i++)
            {
                var row = sheet.Cells[i, 1, i, sheet.Dimension.End.Column];

                ModelNomenclature nomenclature = new();

                nomenclature.Ord = int.Parse(row[i, 1].Value.ToString());
                nomenclature.Kod = int.Parse(row[i, 2].Value.ToString());
                nomenclature.Name = row[i, 2].Value.ToString();

                modelNomenclatures.Add(nomenclature);
            }

            return modelNomenclatures;
        }
        public static byte[] CreateExcelResult(IEnumerable<ModelResultBullet> modelResultBullets)
        {
            ExcelPackage excelPackage = new();
            ExcelWorksheet excelWorkSheet = excelPackage.Workbook.Worksheets.Add("Лист 1");

            excelWorkSheet.Column(1).Width = 50;
            excelWorkSheet.Row(1).Height = 20;

            excelWorkSheet.Cells[1, 1, 2, 1].Merge = true;
            excelWorkSheet.Cells[1, 1, 2, 1].Value = "НАИМЕНОВАНИЕ";
            excelWorkSheet.Cells[1, 1, 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[1, 1, 2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            excelWorkSheet.Cells[1, 2, 1, 5].Merge = true;
            excelWorkSheet.Cells[1, 2, 1, 5].Value = "Приволжский федеральный округ";
            excelWorkSheet.Cells[1, 2, 1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[1, 2, 1, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            excelWorkSheet.Cells[1, 6, 1, 11].Merge = true;
            excelWorkSheet.Cells[1, 6, 1, 11].Value = "Уральский федеральный округ";
            excelWorkSheet.Cells[1, 6, 1, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[1, 6, 1, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            excelWorkSheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            excelWorkSheet.Cells[1, 1, 2, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            excelWorkSheet.Cells[1, 1, 2, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            excelWorkSheet.Cells[2, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            excelWorkSheet.Cells[1, 2, 1, 11].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            excelWorkSheet.Cells[1, 2, 1, 11].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            excelWorkSheet.Cells[1, 2, 1, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            excelWorkSheet.Cells[1, 2, 1, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            excelWorkSheet.Cells[2, 2, 2, 11].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            excelWorkSheet.Cells[2, 2, 2, 11].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            excelWorkSheet.Cells[2, 2].Value = "Уфа";
            excelWorkSheet.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 3].Value = "Ижевск";
            excelWorkSheet.Cells[2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 4].Value = "Пермь";
            excelWorkSheet.Cells[2, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 5].Value = "Оренбург";
            excelWorkSheet.Cells[2, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 6].Value = "Курган";
            excelWorkSheet.Cells[2, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 7].Value = "Екатеринбург";
            excelWorkSheet.Cells[2, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 8].Value = "Тюмень";
            excelWorkSheet.Cells[2, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 9].Value = "Ханты-Мансийск";
            excelWorkSheet.Cells[2, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 10].Value = "Салехард";
            excelWorkSheet.Cells[2, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelWorkSheet.Cells[2, 11].Value = "Челябинск";
            excelWorkSheet.Cells[2, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            int numberRow = 3;

            foreach (var rowBullet in modelResultBullets)
            {
                excelWorkSheet.Cells[numberRow, 1].Value = rowBullet.Name;
                excelWorkSheet.Cells[numberRow, 2].Value = rowBullet.Ufa;
                excelWorkSheet.Cells[numberRow, 3].Value = rowBullet.Ijevsk;
                excelWorkSheet.Cells[numberRow, 4].Value = rowBullet.Perm;
                excelWorkSheet.Cells[numberRow, 5].Value = rowBullet.Orenburg;
                excelWorkSheet.Cells[numberRow, 6].Value = rowBullet.Kurgan;
                excelWorkSheet.Cells[numberRow, 7].Value = rowBullet.Ekaterinburg;
                excelWorkSheet.Cells[numberRow, 8].Value = rowBullet.Tumen;
                excelWorkSheet.Cells[numberRow, 9].Value = rowBullet.Hanty;
                excelWorkSheet.Cells[numberRow, 10].Value = rowBullet.Salehard;
                excelWorkSheet.Cells[numberRow, 11].Value = rowBullet.Chelyabinsk;
                numberRow++;
            }

            return excelPackage.GetAsByteArray();
        }
        public static void SaveFile(byte[] dataFiles, string path)
        {
            try
            {
                File.WriteAllBytes(path, dataFiles);

                Console.WriteLine($"\nСохранен файл: {Path.GetFileName(path)} в директорию: {path}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка: {ex.Message + ex.StackTrace}");
            }
        }
        public static string CreateDirectory(string pathDirectiry)
        {
            if (!Directory.Exists(pathDirectiry))
            {
                Directory.CreateDirectory(pathDirectiry);
                return $"Создана директория: {pathDirectiry}";
            }

            return $"Директория существует: {pathDirectiry}";
        }
    }
}
