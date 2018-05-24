using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Debiteringsformaterare
{
    public enum FaktureringsID
    {
        DebiteringsFilensSkapningsDatum = 001,
        BokningsObjekt = 005,
        Summa = 004,
    }
    public class FaktureringsObjekt
    {

        public FaktureringsID Id;
        public string Lagenhet;
        public string Datum;
        public string Namn;
        public string Typ;
        public string Bokning;
        public float Kostnad;
        public string Fran_Datum;
        public string Till_Datum;

        public FaktureringsObjekt()
        {

        }

        public FaktureringsObjekt(FaktureringsID _id, string _lagenhet, string _datum, string _namn, string _typ, string _bokning, float _kostnad)
        {
            
            this.Id = _id;
            this.Lagenhet = _lagenhet;
            this.Datum = _datum;
            this.Namn = _namn;
            this.Typ = _typ;
            this.Bokning = _bokning;
            this.Kostnad = _kostnad;
        }

        public FaktureringsObjekt(FaktureringsID _id, string _lagenhet, float _kostnad)
        {
            this.Id = _id;
            this.Lagenhet = _lagenhet;
            this.Kostnad = _kostnad;
        }

        public override string ToString()
        {
            return base.ToString();
        }

    }
}
