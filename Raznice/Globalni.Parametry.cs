using System;
using System.Web;
using System.Xml;
using System.Collections.Specialized;


namespace Raznice.Globalni
{
	/// <summary>
	/// Tøída obsahuje všeobecné parametry webové aplikace naèítané z XML souboru Parametry.xml
	/// Aby bylo docíleno co moná nejvyšší aktuálnosti, zjišuje se pøi kadém poadavku
	/// na hodnotu parametru aktuálnost drenıch dat (zjištìní data souboru). 
	/// </summary>
	public class Parametry
	{
		#region Veøejné statické vlastnosti


        /// <summary>
        /// jmeno aplikace
        /// </summary>
        public static string aplikace
        {
            get
            {
                return DejHodnotuParametru("aplikace");
            }
        }

        /// <summary>
        /// z jakeho pole ae pouzije konstrukce pro tisk stitku
        /// </summary>

        public static bool tisk_z_pole_prijmeni
        {
            get
            {
                return (Int32.Parse(DejHodnotuParametru("tisk_z_pole_prijmeni")) == 1 ? true : false) ;
            }
        }
        #endregion


        #region Neveøejnı kód
        #region Vlastnosti
        /// <summary>
		/// Urèuje relativní cestu k XML souboru s parametry
		/// </summary>
        private const string PARAMETRY_XML_RELATIVNI_CESTA = "ParametryAplikace.xml";
		/// <summary>
		/// Uchovává seznam parametrù naètenı z XML dokumentu
		/// </summary>
		private static HybridDictionary parametr;
		/// <summary>
		/// Uchovává datum modifikace naposledy naèteného XML souboru 
		/// </summary>
		private static DateTime modifikovano;
		/// <summary>
		/// Uchovává automaticky generovanou absolutní cestu k XML souboru s parametry
		/// </summary>
		private static string absolutniCestaXmlDokumentu;
		#endregion

		#region Konstruktory
		/// <summary>
		/// Statickı konstruktor, kterı se volá automaticky pøi prvním poadavku
		/// na jakoukoli statickou vlastnost tøídy.
		/// </summary>
		static Parametry()
		{
			// Pøevádím relativní cestu ke XML dokumentu danou konstantou PARAMETRY_XML_RELATIVNI_CESTA
			// na cestu absolutni
			Parametry.absolutniCestaXmlDokumentu = 
				Globalni.Nastroje.DejCestuAplikace() + PARAMETRY_XML_RELATIVNI_CESTA;
//            Parametry.absolutniCestaXmlDokumentu =
//                HttpContext.Current.Request.MapPath(PARAMETRY_XML_RELATIVNI_CESTA);

			// Vytváøím novı seznam parametrù
			Parametry.parametr = new HybridDictionary();
			// Nastavuji datum poslední modifikace parametrù
			Parametry.modifikovano = DateTime.MinValue;

			// Provádím aktualizaci parametrù
			Aktualizuj();
		}
		#endregion

		#region Metody
		/// <summary>
		/// Vrací hodnotu poadovaného parametru.
		/// </summary>
		/// <param name="nazevParametru">Název poadovaného parametru</param>
		/// <returns>Øetìzec s hodnotou parametru</returns>
		private static string DejHodnotuParametru(string nazevParametru)
		{
			// Provedu aktualizaci parametrù
			// V pøípadì, e se ukáe, e tím dochází k vıraznému zpomalení aplikace,
			// lze automatickou aktualizaci vypnout zakomentováním následujícího øáadku.
			// Tím bude k aktualizaci docházet jen pøi inicializaci této tøídy (v konstruktoru)
			// nebo v pøípadì zavolání statické metody Aktualizuj();
			Parametry.Aktualizuj();
			// Zjistím zda seznam parametrù obsahuje poadovanı parametr. Paklie ano,
			// vracím jeho hodnotu. Paklie ne, vrátím prázdnı øetìzec.
			return Parametry.parametr.Contains(nazevParametru) ?
				Parametry.parametr[nazevParametru].ToString() : String.Empty;
		}

		/// <summary>
		/// Aktualizuje seznam parametrù
		/// </summary>
		public static void Aktualizuj()
		{
			// Uzamykám parametry proti zmìnám v jiném vláknu
			lock(Parametry.parametr)
			{
				// Naèítám poslední datum zmìny parametrického XML souboru
				DateTime zapsano = new System.IO.FileInfo(Parametry.absolutniCestaXmlDokumentu).LastWriteTime;
				// Zjisuji zda došlo od poslední aktualizace ke zmìnì data - existuje novìjší verze
				if (zapsano > Parametry.modifikovano)
				{
					// Naètu novou verzi parametrù a poznaèím si datum a èas zmìny
					Parametry.Nacti();
					Parametry.modifikovano = zapsano;
				}
			}
		}

		/// <summary>
		/// Naèítá parametry z XML souboru
		/// </summary>
		private static void Nacti() 
		{
			// Vyprazòuji seznam parametrù
			Parametry.parametr.Clear();

			// Naèítám XML dokument obsahující seznam parametrù
			XmlDocument xmlDokument = new XmlDocument();
			xmlDokument.Load(absolutniCestaXmlDokumentu);
			
			// Naèítám elementy se jménem "parametr"
			XmlNodeList parametry = xmlDokument.GetElementsByTagName("parametr");
			foreach(XmlNode par in parametry)
			{
				// Pøidávám jednotlivé parametry do seznamu parametrù
				Parametry.parametr.Add(
					par.Attributes["nazev"].Value,
					par.Attributes["hodnota"].Value);
			}
		}
		#endregion
		#endregion
	}
}
