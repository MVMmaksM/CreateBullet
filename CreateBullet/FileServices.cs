using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateBullet
{
    internal class FileServices
    {
        public string? GetFileName(string pathFile)
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
        public void ExcelToDataTable(DataTable dataTable, string? pathExcelFile)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;               

                ExcelPackage package = new ExcelPackage(pathExcelFile);
                ExcelWorksheet sheet = package.Workbook.Worksheets[0];                

                for (int i = 1; i <= sheet.Dimension.End.Row; i++)
                {
                    var row = sheet.Cells[i, 1, i, sheet.Dimension.End.Column];
                    DataRow newRow = dataTable.NewRow();                             

                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Text;
                    }

                    dataTable.Rows.Add(newRow);                    
                }

                Console.WriteLine($"\nСчитано {dataTable.Rows.Count} строк из файла: {Path.GetFileName(pathExcelFile)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка: {ex.Message + ex.StackTrace}");
            }
        }

        public void SaveFile(ExcelPackage? excelPackage, string path)
        {
            try
            {
                byte[]? fileBin = excelPackage?.GetAsByteArray();

                File.WriteAllBytes(path, fileBin);

                Console.WriteLine($"\nСохранен файл: {Path.GetFileName(path)} в директорию: {path}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nОшибка: {ex.Message + ex.StackTrace}");
            }           
        }

        public void CreateDirectory(params string[] pathDirectory) 
        {
            foreach (var path in pathDirectory )
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    Console.WriteLine($"\nСоздана директория: {path}");
                }
            }                                
        }
    }
}
