using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulletCreate.Model
{
    internal class ModelNameColumnDataMarts
    {
        public string? NameColumn { get; set; }
        public byte NumberColumn { get; set; }
        public static IEnumerable<ModelNameColumnDataMarts> CreateNameColumnDataMarts()
        {           
            return new List<ModelNameColumnDataMarts>()
            {
                new ModelNameColumnDataMarts {NameColumn = "Наименование", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Код", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Уфа", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Ижевск", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Пермь", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Оренбург", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Курган", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Екатеринбург", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Тюмень", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Ханты-Мансийск", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Салехард", NumberColumn = 0 },
                new ModelNameColumnDataMarts {NameColumn = "Челябинск", NumberColumn = 0 },
            };
        }
    }
}
