using System.Data;

namespace CreateBullet
{
    internal class Program
    {
        static string pathListKor = $"{Environment.CurrentDirectory}\\Короткий";
        static string pathListPoln = $"{Environment.CurrentDirectory}\\Полный";
        static string pathListVigr = $"{Environment.CurrentDirectory}\\Выгрузка";
        static string pathListBull = $"{Environment.CurrentDirectory}\\Бюллетень";

        static void Main(string[] args)
        {
            Console.WriteLine("\n--------------Старт------------");

            FileServices fileServices = new();
            fileServices.CreateDirectory(pathListKor, pathListPoln, pathListVigr, pathListBull);

            DataTable listKor = new DataTable();
            listKor.Columns.AddRange(new DataColumn[3]
           { new DataColumn("porydok", typeof(int)),
            new DataColumn("kod", typeof(int)),
            new DataColumn ("name", typeof(string)),
           });

            DataTable listPolny = new DataTable();
            listPolny.Columns.AddRange(new DataColumn[3]
           { new DataColumn("porydok", typeof(int)),
            new DataColumn("kod", typeof(int)),
            new DataColumn ("name", typeof(string)),
           });

            DataTable listVigr = new DataTable();
            listVigr.Columns.AddRange(new DataColumn[13]
           { new DataColumn("name", typeof(string)),
            new DataColumn("kod", typeof(int)),
            new DataColumn ("ufa", typeof(string)),
            new DataColumn ("ijevsk", typeof(string)),
            new DataColumn ("perm", typeof(string)),
            new DataColumn ("orenburg", typeof(string)),
            new DataColumn ("kurgan", typeof(string)),
            new DataColumn ("ekaterinburg", typeof(string)),
            new DataColumn ("tumen", typeof(string)),
            new DataColumn ("hanty", typeof(string)),
            new DataColumn ("salehard", typeof(string)),
            new DataColumn ("chelyabinsk", typeof(string)),
            new DataColumn ("chelyabinsk2", typeof(string)),
           });
            
            if (fileServices.GetFileName(pathListKor) != null && fileServices.GetFileName(pathListPoln) != null && fileServices.GetFileName(pathListVigr) != null)
            {
                fileServices.ExcelToDataTable(listKor, fileServices.GetFileName(pathListKor));
                fileServices.ExcelToDataTable(listPolny, fileServices.GetFileName(pathListPoln));
                fileServices.ExcelToDataTable(listVigr, fileServices.GetFileName(pathListVigr));

                Bulletines bulletines = new();
                fileServices.SaveFile(bulletines.Create(listPolny, listVigr), pathListBull + "\\Для бюллетеня полный.xlsx");
                fileServices.SaveFile(bulletines.Create(listKor, listVigr), pathListBull + "\\Для бюллетеня короткий.xlsx");

                Console.WriteLine("\n----------------Завершено--------------");
                Console.ReadLine();
            }

            Console.ReadLine();
        }
    }
}