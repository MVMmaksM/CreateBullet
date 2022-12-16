using BulletCreate.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulletCreate.Bullet
{
    internal class Bulletin
    {
        public static IEnumerable<ModelResultBullet> UniteForBullet(IEnumerable<ModelDataMarts> dataMarts, IEnumerable<ModelNomenclature> nomenclatures)
        {
            var resultUniteForBullet = dataMarts.Join(nomenclatures, d => d.Kod, n => n.Kod,
                (d, n) => new ModelResultBullet
                {
                    Ord = n.Ord,
                    Name = n.Name,
                    Ufa = d.Ufa,
                    Ijevsk = d.Ijevsk,
                    Perm = d.Perm,
                    Orenburg = d.Orenburg,
                    Kurgan = d.Kurgan,
                    Ekaterinburg = d.Ekaterinburg,
                    Tumen = d.Tumen,
                    Hanty = d.Hanty,
                    Salehard = d.Salehard,
                    Chelyabinsk= d.Chelyabinsk
                }).OrderBy(n=>n.Ord);

            return resultUniteForBullet;
        }
    }
}
