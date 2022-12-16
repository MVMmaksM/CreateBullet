using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BulletCreate.Model;
using OfficeOpenXml;

namespace BulletCreate.Services
{
    internal class FileServices
    {
        private static string? GetFileName(string pathFile)
        {
            try
            {
                string[] files = Directory.GetFiles(pathFile);

                if (files.Length > 1)
                {
                    Console.WriteLine("\nДолжен быть один файл!");
                    return null;
                }
                else if (files.Length == 0)
                {
                    Console.WriteLine("\nНет файла с порядком позиций!");
                    return null;
                }

                return files[0];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка: {ex.Message + ex.StackTrace}");
                return null;
            }
        }
        public static IEnumerable<ModelDataMarts> LoadInDataMarts(string pathFile)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(GetFileName(pathFile));
            ExcelWorksheet sheet = package.Workbook.Worksheets[0];
            
            IEnumerable<ModelDataMarts> dataMarts = new List<ModelDataMarts>();

            for (int i = 1; i <= sheet.Dimension.End.Row; i++)
            {
                var row = sheet.Cells[i, 1, i, sheet.Dimension.End.Column];

                ModelDataMarts rowDatatMarts = new ();

                rowDatatMarts.Name = row[i, 1].Value.ToString();
                rowDatatMarts.Kod = int.Parse(row[i, 2].Value.ToString());
                rowDatatMarts.
            }
        }
    }
}
