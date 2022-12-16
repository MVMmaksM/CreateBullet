using BulletCreate.Bullet;
using BulletCreate.Model;
using BulletCreate.Services;

namespace BulletCreate
{
    internal class Program
    {
        private static string _pathDirectoryDataMarts = $"{Environment.CurrentDirectory}\\Выгрузка из витрины данных";
        private static string _pathDirectoryNomenclaturePoln = $"{Environment.CurrentDirectory}\\Полный перечень";
        private static string _pathDirectoryNomenclatureKor = $"{Environment.CurrentDirectory}\\Короткий перечень";
        private static string _pathDirectorySaveBullet = $"{Environment.CurrentDirectory}\\Бюллетень";
        private static List<ModelDataMarts>? _dataMarts;
        private static List<ModelNomenclature>? _nomenclaturePoln;
        private static List<ModelNomenclature>? _nomenclatureKor;

        static void Main(string[] args)
        {
            Console.WriteLine("---------Старт------------");
            Console.WriteLine("\n--------------Проверка директорий----------");

            Console.WriteLine();

            Console.WriteLine(FileServices.CreateDirectory(_pathDirectoryDataMarts));
            Console.WriteLine(FileServices.CreateDirectory(_pathDirectoryNomenclaturePoln));
            Console.WriteLine(FileServices.CreateDirectory(_pathDirectoryNomenclatureKor));
            Console.WriteLine(FileServices.CreateDirectory(_pathDirectorySaveBullet));

            Console.WriteLine("\n------------Загрузка данных---------------");

            _dataMarts = (List<ModelDataMarts>)FileServices.LoadInDataMarts(_pathDirectoryDataMarts);
            _nomenclaturePoln = (List<ModelNomenclature>)FileServices.LoadInNomenclature(_pathDirectoryNomenclaturePoln);
            _nomenclatureKor = (List<ModelNomenclature>)FileServices.LoadInNomenclature(_pathDirectoryNomenclatureKor);

            if (_dataMarts is not null && _nomenclaturePoln is not null && _nomenclaturePoln is not null)
            {
                Console.WriteLine($"Загружено из витрин данных: {_dataMarts.Count} записей");
                Console.WriteLine($"Загружено из полного перечня: {_nomenclaturePoln.Count} записей");
                Console.WriteLine($"Загружено из короткого перечня: {_nomenclatureKor.Count} записей");

                Console.WriteLine("\n----------Создание бюллетеней-------------");
                FileServices.SaveFile(FileServices.CreateExcelResult(Bulletin.UniteForBullet(_dataMarts, _nomenclatureKor)),_pathDirectorySaveBullet+"\\Для бюллетеня короткий.xlsx");
                FileServices.SaveFile(FileServices.CreateExcelResult(Bulletin.UniteForBullet(_dataMarts, _nomenclaturePoln)),_pathDirectorySaveBullet+"\\Для бюллетеня полный.xlsx");

                Console.WriteLine("\n----------Выполнено----------------");
            }

            Console.ReadLine();
        }
    }
}