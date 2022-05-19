using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.OleDb;

using System.Runtime.InteropServices;

namespace Raznice
{
    
    public partial class frmRaznice : Form
    {
        /// <summary>
        /// počet vyražených dozimetrů
        /// </summary>
        int DozCount;               // počet vyražených dozimetrů

        /// <summary>
        /// počet dozimetrů, které se mají orazit
        /// </summary>
        int DozMaxCount;            // počet dozimetrů, které se mají orazit        

        /// <summary>
        /// pole čísel dozimetrů
        /// </summary>
        string[] DozStr;            // pole čísel dozimetrů

        /// <summary>
        /// indikace druhu ražení
        /// </summary>
        bool DozFile;               // indikace druhu ražení

        /// <summary>
        /// od ktere pozice - radku v souboru se ma razit
        /// </summary>
        int DozPozice;              // od ktere pozice - radku v souboru se ma razit


        /// <summary>
        /// jmeno dbf souboru s razicim planem
        /// </summary>
        string dbFileName = "";     // jmeno dbf souboru s razicim planem

        frmRaznice formRaz;         // nastaveny object formu ve Form_Load pro parametr zasilani LogMessage
        bool tisk_z_pole_prijmeni;  // jake pole (prijmeni nebo Tisk_2) z tabulky se pouzije pro tisk stitku

        int DozNum;                 // číslo dozimetru
        int DozVyrazeno;            // pocitadlo vyrazenych doz v pripade razeni od - do

        /// <summary>
        /// je proces razeni?
        /// </summary>
        bool vProcesuRazeni = false;

        /// <summary>
        /// kolikrat probehne pokus o nastaveni textu na raznici
        /// </summary>
        int koleckoSendTextKolikrat = 0;
        /// <summary>
        /// kolik sec se pocka, po zavolani Reset(), pri neuspechu zapisu textu na dozimetr
        /// </summary>
        int cekejPoReset = 0;
        /// <summary>
        /// kolikrat probehne pokus o zjisteni konce tisku dozimetru raznici
        /// </summary>
        int koleckoFinishKolikrat = 0;
        /// <summary>
        /// kolik sec se pocka, nez se zavola dalsi pokus cteni ReadFinishOK()
        /// </summary>
        int cekejPredReadFinishOK = 0;
        /// <summary>
        /// kolik sec se pocka, po zavolani Start(), pri zapisu textu na dozimetr
        /// </summary>
        int cekejPoStart = 0;


        private class Item
        {
            public string Name;
            public int Value;
            public Item(string name, int value)
            {
                Name = name; Value = value;
            }
            public override string ToString()
            {
                // Generates the text shown in the combo box
                return Name;
            }
        }

        public string[] LoadFile(string fileName)
        {
            
            try
            {
                StreamReader file = new StreamReader(fileName);
                char[] separator = new char[] { '\n' };
                string[] res = file.ReadToEnd().Split(separator, StringSplitOptions.RemoveEmptyEntries);
                //Close the file
                file.Close();

                Globalni.Nastroje.LogMessage("Natazen soubor: " + fileName, false, "Information", formRaz);
                return res;
            }
            catch { return null; }
        }

        private static string DecodeFromUtf8(string utf8_String)
        {
            //string utf8_String = "dayâ€™s";
            byte[] bytes = Encoding.Default.GetBytes(utf8_String);
            

            //utf8_String = Encoding.UTF8.GetString(bytes);
            utf8_String = Encoding.GetEncoding("windows-1250").GetString(bytes);
            
            return utf8_String;
        }

        private string DejSarziFilmu()
        {
            //1A Vachata
            string strSarze = "";
            strSarze = txtSarze.Text.ToString().ToUpper();
            if (strSarze == String.Empty)
                strSarze = "A";
            return strSarze;
        }

        /// <summary>
        /// Podle pozice v parametru vypise na form cislo doz, popis doz, EAN
        /// </summary>
        /// <param name="index"> index pole souboru</param>
        private void RozeberDozStr(int index)
        {

            try
            {
                // 05019017;1A Vachata
                string[] rowArr = DozStr[index].Split(';');

                // 05019017
                //lblDozNum.Text = rowArr[0].Trim('"', ' ');

                // 10800427;3 Kozloduy_427
                // 10168004427;3 Kozloduy_427
                string pom = "";
                pom = rowArr[0].Trim('"', ' '); //10168004427 --> 10rr800o427

                lblDozNum.Text = pom.Substring(0, 2) +      //10
                                 pom.Substring(4, 3) +      //800
                                 pom.Substring(8, 3) ;     //427

                /*
                //1 Vachata
                //lblDozPopis.Text = DecodeFromUtf8(rowArr[1].Trim('"', ' '));
                lblDozPopis.Text = Decodecharset(rowArr[1].Trim('"', ' '));
                //1C Vachata
                lblDozPopis.Text = lblDozPopis.Text.Substring(0, 1) +
                                    DejSarziFilmu() +
                                    lblDozPopis.Text.Substring(1, lblDozPopis.Text.Length-1);
                */
                string pom1 = "";

                pom1 = Decodecharset(rowArr[1].Trim('"', ' '));
                lblDozPopis.Text = pom1.Substring(0, 1) +
                                    DejSarziFilmu() + "_" + // 3A_
                                    pom.Substring(0, 2) + "_" +  //  10_
                                    //pom.Substring(2, 2) + "_" +  //  16_
                                    pom.Substring(4, 3) + "/" +  //  800/
                                    pom.Substring(7, 1) + "_" +  //  4
                                    pom.Substring(8, 3) +   //  427
                                    pom1.Substring(1, pom1.Length - 1); // Vachata


                lblDozPopisEAN.Text = pom1.Substring(0, 1) + //3
                                    pom.Substring(0, 2) +    //10
                                    pom.Substring(2, 2) +    //16
                                    pom.Substring(4, 3) +    //800
                                    pom.Substring(7, 1) +    //4
                                    pom.Substring(8, 3) ;    //427
                                    

            }
            catch (Exception e)
            {
                lblDozNum.Text = "";
                lblDozPopis.Text = "";
                lblDozPopisEAN.Text = "";
            }
        }

        private bool JeTamCisloDozimetru(string cisloDozimetru)
        {
            bool JeTam = false;
            DozPozice = 0;

            int i = 0;
            while (i < DozStr.Length)
            {

                // 05019017;1A Vachata
                string[] rowArr = DozStr[i].Split(';');

                // 05019017
                string DozNum = rowArr[0].Trim('"', ' ');

                if (cisloDozimetru == rowArr[0].Trim('"', ' '))
                {
                    if ((txtRazitDoz.Text.ToString().Replace(" ", "") == string.Empty) || (int.Parse(txtRazitDoz.Text.ToString().Replace(" ", "")) == 0))
                    {
                        MessageBox.Show("Hodnota 'Počet dozimetrů' není zadána.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    // poznamenam si pozici
                    DozPozice = i+1;
                    return true;
                }

                i++;
                if (i > DozStr.Length)
                {
                    JeTam = false;
                }


            }

            if (JeTam == false)
                MessageBox.Show("Číslo dozimetru v souboru nebylo nalezeno", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error); 
            return JeTam;
        }



        /// <summary>
        /// Vyrazeni dozimetru ze souboru TAB3 
        /// </summary>
        public void StartNV2()
        {
            Vlastnosti.popisStavuRaznice popisStavuRaznice = new Vlastnosti.popisStavuRaznice();

            #region Kontrola
            if (txtSarze.Text.Replace(" ", "") == String.Empty)
            {
                MessageBox.Show("Šarže filmu není zadána", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("StartNV2(), Šarže filmu není zadána", false, "Error", formRaz);
                return;
            }
            if (txtTyp.Text == String.Empty)
            {
                MessageBox.Show("Typ filmu není zadán", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return;

            }
            if (txtTyp.Text != "1" && txtTyp.Text != "2" && txtTyp.Text != "3")
            {
                MessageBox.Show("Typ filmu není zadán v intervalu 1 - 3", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return;

            }
            int nTyp = 0;
            if (!int.TryParse(txtTyp.Text, out nTyp))
            {
                MessageBox.Show("Typ filmu není zadán korektně v intervalu 1 - 3", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return;

            }
            #endregion

            bool ok = Init();
            if (!ok)
            {
                popisStavuRaznice = DejPopisStavu();

                lblStatus.Text = "Chyba komunikace: " + popisStavuRaznice.stavText.ToString();
                Globalni.Nastroje.LogMessage("StartNV2() Init, Chyba komunikace: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                chkReady.Checked = false;
            }
            chkReady.Checked = true;

            popisStavuRaznice = DejPopisStavu();
            if (popisStavuRaznice.nStatusId != 3) //chyba, řízení vypnuto
            {
                MessageBox.Show("Raznice není připravena: " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("StartNV2() Init, Raznice není připravena: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                return;
            }

            //string txt = "";
            DozCount = 0;
            DozVyrazeno = 0;
            lblCount2.Text = "0";
            bool konec = false;
            int err = 0;
            // razba ze souboru TAB3 a nebo bez TAB2
            if (DozFile)
            {
                // ze souboru TAB3
                DozStr = LoadFile(txtFile.Text);    // 05019017;1 Vachata

                if (!(DozStr == null))
                {
                    int locDozPozice = 0;

                    // vynulovat vstupy
                    lblCount2.Text = "0";


                    // tiskne se vse nebo jenom podmnozina 
                    if (txtRazitOdDoz.Text.Trim().Length > 0)
                        if (JeTamCisloDozimetru(txtRazitOdDoz.Text.Replace(" ", "").Trim()))
                        {
                            locDozPozice = DozPozice - 1;
                            if (locDozPozice < 0)
                                locDozPozice = 0;
                        }
                        else
                            return;


                    RozeberDozStr(locDozPozice);
                    DozMaxCount = DozStr.Length;

                    if (lblDozNum.Text == String.Empty || lblDozPopis.Text == String.Empty || lblDozPopisEAN.Text == String.Empty)
                    {
                        MessageBox.Show("Nelze řádně dekódovat obsah ze souboru, špatný soubor?", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Globalni.Nastroje.LogMessage("StartNV2(), Nelze řádně dekódovat obsah ze souboru, špatný soubor?: " + txtFile.Text.ToString(), false, "Error", formRaz);
                        return;
                    }

                }
                else
                {
                    MessageBox.Show("Nelze načíst soubor", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Globalni.Nastroje.LogMessage("StartNV2(), Nelze načíst soubor: " + txtFile.Text.ToString(), false, "Error", formRaz);
                    return;
                }
            }

            // prvni doz, dalsi se resi pres timer2 ...
            #region nepodstatne
            //            if (DozFile)
            //            {
            //                // razeni pres soubor TAB3
            //                bool vysledek = false;
            //                bool jaktisk = false;

            //                // tiskne se vse nebo jenom podmnozina ?
            //                if ((txtRazitOdDoz.Text.Replace(" ", "").Trim().Length > 0))
            //                {
            //                    // podmnozina
            //                    // tisknu az z timeru2
            //                    System.Threading.Thread.Sleep(1000);
            ////                    timer2.Enabled = true;
            //                }
            //                else
            //                {
            //                    // tiskne se vse ze souboru
            //                    // ponovu se nastavi co tisknout a posle se na razbu kde se zaroven i tiskne
            //                    #region razbaV2 s tiskem dozimetru NEW

            //                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
            //                    popisStavuRaznice = DejPopisStavu();
            //                    if ((popisStavuRaznice.nStatusId != 3)) //chyba, řízení vypnuto
            //                    {
            //                        MessageBox.Show("StartN(): " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //                        Globalni.Nastroje.LogMessage("StartN(): " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
            //                        vysledek = false;
            //                    }

            //                    if (chkRazitDozimetry.Checked == true)
            //                    {
            //                        Globalni.Nastroje.LogMessage("StartN(), StartText(txt, txt.Length): " + txt.ToString(), false, "Error", formRaz);


            //                        //string numZdroj = lblDozNum.Text.ToString().Trim();
            //                        string nameZdroj = lblDozPopis.Text.ToString().Trim();
            //                        string nameZdrojEAN = lblDozPopisEAN.Text.ToString().Trim();

            //                        vysledek = NaRazitDozV2(txt_numDoz: txt, txt_nameZdroj: lblDozPopis.Text.ToString().Trim(), txt_numZdroj: lblDozNum.Text.ToString().Trim(), txtTyp.Text.ToString());



            //                    }
            //                    else
            //                        vysledek = true;

            //                    // tisk dalsich dozimetru pres timer
            //                    if (vysledek == true)
            //                    {
            //                        System.Threading.Thread.Sleep(1000);
            // //                       timer2.Enabled = true;
            //                    }
            //                    #endregion
            //                }
            //            }
            #endregion

            vProcesuRazeni = true;
            EnablingDone(false);
            while (!konec)
            {
                // v pripade, ze se zmackne STOP 
                if (!vProcesuRazeni)
                {
                    konec = true;
                    break;
                }

// --- timer
                DozCount += 1;

                // tiskne se vse nebo jenom podmnozina ?
                if ((txtRazitOdDoz.Text.Replace(" ", "").Trim().Length > 0)
                    &&
                        (!((DozCount >= DozPozice)
                        &&
                        //(DozCount <= (DozPozice + int.Parse(txtRazitDoz.Text.Replace(" ", "").Trim()) - 1))))
                        //(DozCount <= (DozPozice + int.Parse(txtRazitDoz.Text.Replace(" ", "").Trim()) + 1))))
                        (DozCount <= (DozPozice + int.Parse(txtRazitDoz.Text.Replace(" ", "").Trim())))))
                    )
                {
                    if ((err > 0) || (DozCount >= DozMaxCount))
                    {
                        // timer2.Enabled = false;
                        konec = true;
                        //return;
                        vProcesuRazeni = false;
                        break;
                    }
                    // vynechavam razeni, nejsem v intervalu
                    lblStatus.Text = "Skip dozimetru";
                    Globalni.Nastroje.LogMessage("StartNV2(), Skip dozimetru", false, "Information", formRaz);
                    //return;
                    continue;
                }
                else
                {
                    DozVyrazeno += 1;
                }

                lblCount2.Text = DozCount.ToString();


                if (DozCount < DozStr.Length)
                {
                    RozeberDozStr(index: DozCount);
                }
                else
                {
                    RozeberDozStr(index: 0);
                }

                if ((err > 0) || (DozCount >= DozMaxCount))
                {
                    //                timer2.Enabled = false;
                    //return;
                    konec = true;
                    vProcesuRazeni = false;
                    break;
                }


                //RozeberDozStr(index: DozCount);
                string numZdroj = lblDozNum.Text.ToString().Trim();
                string nameZdroj = lblDozPopis.Text.ToString().Trim();
                string nameZdrojEAN = lblDozPopisEAN.Text.ToString().Trim();

                //pokud je zadano omezeni intervalu dozimetru k tisku

                int i = 0;
                bool vysledek = false;
                // priznak, ze se ma vubec provadet razeni
                if (chkRazitDozimetry.Checked == true)
                {
                    Globalni.Nastroje.LogMessage("StartNV2() NaRazitDozV2 txt_numDoz: " + numZdroj.ToString(), false, "Information", formRaz);

                    vysledek = NaRazitDozV2(txt_numDoz: numZdroj, txt_nameZdroj: nameZdroj, txt_numZdroj: nameZdrojEAN, typeDoz: txtTyp.Text.ToString());

                }
                else
                    vysledek = true;

                if (!vysledek)
                {
                    konec = true;
                    vProcesuRazeni = false;
                    break;
                }

            } // end while (konec)

            vProcesuRazeni = false;
            DozFile = false;
            EnablingDone(true);


            // --- timer end


        }


        #region ImportDLL

#if DLL
   
        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool Init();

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool Ping();

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool Start();

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool Reset();

        //////////////////////

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool ReadStatus(ref short nStatus);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool ReadInfo(ref short nInfo);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool ReadError(ref short nError);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool ReadFinishOK(ref bool lOK);

        //////////////////////
        //static extern bool SendType(int nType);
        //static extern bool SendType(char nType);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendType([MarshalAs(UnmanagedType.LPWStr)] string txt);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendTextName([MarshalAs(UnmanagedType.LPWStr)] string txt, int nLen);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendTextPersonalNo([MarshalAs(UnmanagedType.LPWStr)] string txt, int nLen);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendTextBarCode([MarshalAs(UnmanagedType.LPWStr)] string txt, int nLen);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendTextRazNo([MarshalAs(UnmanagedType.LPWStr)] string txt, int nLen);
#else

        //simulace fci z Raznice.dll        

        /// <summary>
        /// Fce Init() zřizuje konektivitu s PLC. Pokud je návratová hodnota TRUE, tak můžeme volat ostatní fce a to
        /// jednak pro čtení Read a jednak pro zápis Send(proměnná typu PChar je pointer na String, proměnná nLen je délka Stringu).
        /// </summary>
        /// <returns></returns>
        public bool Init()
        {
            return true;
        }

        public bool Ping()
        {
            return true;
        }

        public bool Start()
        {
            return true;
        }

        public bool Reset()
        {
            return true;
        }
        ////////////////////////////        
        public bool ReadStatus(ref short nStatus)
        {
            //nStatus = 3;
            Item itm = (Item)cbStatut.SelectedItem;
            //int selectedIndex = cbStatut.SelectedIndex;
            //cbStatut.Items[selectedIndex];
            nStatus = (short)itm.Value;

            return true;
        }

        public bool ReadInfo(ref short nInfo)
        {
            //nInfo = 2;
            Item itm = (Item)cbInfo.SelectedItem;
            int selectedIndex = cbInfo.SelectedIndex;
            nInfo = (short)itm.Value;

            return true;
        }

        public bool ReadError(ref short nError)
        {
            //nError = 0;
            Item itm = (Item)cbError.SelectedItem;
            nError = (short)itm.Value;

            return true;
        }
        public bool ReadFinishOK(ref bool lOK)
        {
            //lOK = true;
            Item itm = (Item)cbFinishOK.SelectedItem;
            lOK = (itm.Value == 1);

            return true;
        }
        ////////////////////////////

        public bool SendType([MarshalAs(UnmanagedType.LPWStr)] string typ)
        {
            return true;
        }

        public bool SendTextName([MarshalAs(UnmanagedType.LPWStr)] string txt, int nLen)
        {
            //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            Globalni.Nastroje.LogMessage("SendTextName:" + txt.Trim(), false, "Information", formRaz);
            //return true;
            Random random = new Random();
            double r = random.NextDouble();
            //int a = (int)r;
            return (r < 0.5 ? false : true);
        }

        public bool SendTextPersonalNo([MarshalAs(UnmanagedType.LPWStr)] string txt, int nLen)
        {
            //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            Globalni.Nastroje.LogMessage("SendPersonalNo:" + txt.Trim(), false, "Information", formRaz);
            //return true;
            Random random = new Random();
            double r = random.NextDouble();
            //int a = (int)r;
            return (r < 0.5 ? false : true);
        }

        public bool SendTextBarCode([MarshalAs(UnmanagedType.LPWStr)] string txt, int nLen)
        {
            //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            Globalni.Nastroje.LogMessage("SendTextBarCode:" + txt.Trim(), false, "Information", formRaz);
            //return true;
            Random random = new Random();
            double r = random.NextDouble();
            //int a = (int)r;
            return (r < 0.5 ? false : true);
        }

        public bool SendTextRazNo([MarshalAs(UnmanagedType.LPWStr)] string txt, int nLen)
        {
            //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            Globalni.Nastroje.LogMessage("SendTextRazNo:" + txt.Trim(), false, "Information", formRaz);
            //return true;
            Random random = new Random();
            double r = random.NextDouble();
            //int a = (int)r;
            return (r < 0.5 ? false : true);
        }


        //#region stare fce, jen aby to nervalo zatim
        ////simulace fci z Raznice.dll

        //public bool IsReady(ref bool Status)
        //{
        //    Status = true;
        //    return true;
        //}

        //public bool IsDone(ref bool done, ref int Err, ref int Mark)
        //{
        //    done = true;
        //    Err = 0;
        //    Mark = 0;
        //    return true;
        //}


        //public bool StartText([MarshalAs(UnmanagedType.LPStr)] string text, int len)
        //{
        //    //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    Globalni.Nastroje.LogMessage("StartText:" + text.Trim(), false, "Information", formRaz);
        //    //return true;
        //    Random random = new Random();
        //    double r = random.NextDouble();
        //    //int a = (int)r;
        //    return (r < 0.5 ? false : true);
        //}

        //public bool Run()
        //{
        //    return true;
        //}

        //public bool Stop()
        //{
        //    // tohle se pouziva 

        //    return true;
        //}

        //public bool SendText([MarshalAs(UnmanagedType.LPStr)] string text, int len)
        //{
        //    //MessageBox.Show("SendText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    Globalni.Nastroje.LogMessage("SendText:" + text.Trim(), false, "Information", formRaz);
        //    return true;
        //}

        //public bool Mask([MarshalAs(UnmanagedType.LPStr)] string text, int len)
        //{
        //    Globalni.Nastroje.LogMessage("Mask:" + text.Trim(), false, "Information", formRaz);
        //    return true;
        //}

        //public bool PrintCode39([MarshalAs(UnmanagedType.LPStr)] string number, int len, string name, int len2)
        //{
        //    return true;
        //}

        //public bool PrintEAN8([MarshalAs(UnmanagedType.LPStr)] string number, int len, string name, int len2)
        //{
        //    //MessageBox.Show("PrintEAN8:" + name.Trim() + ", " + number.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    Globalni.Nastroje.LogMessage("PrintEAN8:" + name.Trim() + ", " + number.Trim(), false, "Information", formRaz);
        //    return true;
        //}

        //public bool PrintEAN13([MarshalAs(UnmanagedType.LPStr)] string number, int len, string name, int len2)
        //{
        //    //MessageBox.Show("PrintEAN13:" + name.Trim() + ", " + number.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    Globalni.Nastroje.LogMessage("PrintEAN13:" + name.Trim() + ", " + number.Trim(), false, "Information", formRaz);
        //    return true;
        //}

        //public bool SetIP(byte IP1, byte IP2, byte IP3, byte IP4)
        //{
        //    byte[] IPs;
        //    IPs = new byte[4];
        //    IPs[0] = IP1;
        //    IPs[1] = IP2;
        //    IPs[2] = IP3;
        //    IPs[3] = IP4;

        //    /*
        //    string strIPs = "";
        //    strIPs = Encoding.GetEncoding("windows-1250").GetString(IPs);
        //    */

        //    Globalni.Nastroje.LogMessage("SetIP: " + IP1.ToString() + "." + IP2.ToString() + "." + IP3.ToString() + "." + IP4.ToString(), false, "Information", formRaz);
        //    return true;
        //}

        //public bool PistonUp()
        //{
        //    return true;
        //}

        //public bool PistonDown()
        //{
        //    return true;
        //}

        //public bool Eject()
        //{
        //    return true;
        //}

        //public bool ClearInput()
        //{
        //    return true;
        //}

        //public void Disconnect()
        //{

        //}
        //#endregion

#endif

        #endregion

        #region Formular

        public frmRaznice()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //object pro zasilani do LogMessage
            formRaz = this;
            // z ceho (jakeho pole) se tvori stitek
            tisk_z_pole_prijmeni = Globalni.Parametry.tisk_z_pole_prijmeni;
            // nastaveni poctu a zdrzeni v kolecku razeni
            koleckoSendTextKolikrat = Globalni.Parametry.koleckoSendTextKolikrat;
            cekejPoReset = Globalni.Parametry.cekejPoReset;
            cekejPoStart = Globalni.Parametry.cekejPoStart;
            koleckoFinishKolikrat = Globalni.Parametry.koleckoFinishKolikrat;
            cekejPredReadFinishOK = Globalni.Parametry.cekejPredReadFinishOK;


            // v pripade storna u loginu
            if (Vlastnosti.exit == true)
            {
                Application.DoEvents();
                Application.Exit();
            }


            // test na Provider=VFPOLEDB.1
            try
            {
                string filepath = Globalni.Nastroje.DejCestuAplikace();
                if (!filepath.EndsWith("\\"))
                    filepath += "\\";
                OleDbConnection yourConnectionHandler = new OleDbConnection(
                    //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                    @"Provider=VFPOLEDB.1;Data Source=" + filepath);
                yourConnectionHandler.Open();
                if (yourConnectionHandler.State == ConnectionState.Open)
                {
                    yourConnectionHandler.Close();
                }
            }
            catch
            {
                MessageBox.Show("Nenalezen Provider=VFPOLEDB.1", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Globalni.Nastroje.LogMessage("Nenalezen Provider=VFPOLEDB.1", false, "Warning", formRaz);
                this.cmdOtevritPlan.Enabled = false;
            }



            try
            {
                this.Text = this.Text + " [" + (Vlastnosti.allowEdit == true ? "Administrator" : "Uživatel")+ "]";

                // jinak se gridview presosa
                dataGridView1.AutoGenerateColumns = false;
                dataGridView2.AutoGenerateColumns = false;

                // Create the ToolTip and associate with the Form container.
                ToolTip toolTip1 = new ToolTip();

                // Set up the delays for the ToolTip.
                toolTip1.AutoPopDelay = 5000;
                toolTip1.InitialDelay = 1000;
                toolTip1.ReshowDelay = 500;
                // Force the ToolTip text to be displayed whether or not the form is active.
                toolTip1.ShowAlways = true;

                // Set up the ToolTip text 
                toolTip1.SetToolTip(this.cmdOtevritPlan, "Otevření razicího plánu pro ražení");
                toolTip1.SetToolTip(this.cmdOznacitVse, "Označení všech dozimetrů jako 'Zpracováno'");
                toolTip1.SetToolTip(this.cmdOdeznacitVse, "Odeznačení všech dozimetrů jako 'Zpracováno'");
                toolTip1.SetToolTip(this.cmdVyrazit, "Vyražení všech dozimetrů neoznačených jako 'Zpracováno'");

                toolTip1.SetToolTip(this.chkPtatSePredRazbou, "Před každým vyražením dozimetru se musí potvrdit jeho vyražení");
                toolTip1.SetToolTip(this.chkRazitDozimetryTab, "Pokud je vybráno, dozimetr se orazí");
                toolTip1.SetToolTip(this.chkTiskSouborTab, "Pokud je vybráno, štítek pro dozimetr se vytiskne");

                toolTip1.SetToolTip(this.lblDozPopisTab, (tisk_z_pole_prijmeni == true ? "Kontrukce pro text dozimetru z pole 'Příjmení'" : "Kontrukce pro štítek dozimetru z pole 'Tisk řádek_2'"));
                toolTip1.SetToolTip(this.lblEANPopis_radek_2, (tisk_z_pole_prijmeni == true ? "Kontrukce pro štítek dozimetru z pole 'Příjmení'" : "Kontrukce pro štítek dozimetru z pole 'Tisk řádek_2'"));
                toolTip1.SetToolTip(this.cmdOznacitVseVyrazit, "Označení všech podniků jako 'Vyrazit'");
                toolTip1.SetToolTip(this.cmdOdeznacitVseVyrazit, "Odeznačení všech podniků jako 'Vyrazit'");

                //typ filmu (1=malý, 2=velký, 3=velký s otočeným tiskem)
                cbTypFilmu.Items.Add(new Item("1 - malý", 1));
                cbTypFilmu.Items.Add(new Item("2 - velký", 2));
                cbTypFilmu.Items.Add(new Item("3 - velký s otočeným tiskem", 3));
                cbTypFilmu.SelectedIndex = 1;
#if DLL
                // zneviditelnim ovladani simulace
                groupBoxSimulace.Visible = false;

                // zviditelnim STOP/RUN
                groupBoxManualOvladani.Visible = true;

#else
                this.Text = this.Text + " Simulace DLL";
                // zviditelnim ovladani simulace
                groupBoxSimulace.Visible = true;

                // pro simulaci navrat stavu atd.
                cbInit.Items.Add(new Item("Ok", 1));
                cbInit.Items.Add(new Item("False", 0));
                cbInit.SelectedIndex = 0;

                cbStatut.Items.Add(new Item("řízení vypnuto", 0));
                cbStatut.Items.Add(new Item("řízení zapnuto", 1));
                cbStatut.Items.Add(new Item("automatika zapnuta", 2));
                cbStatut.Items.Add(new Item("automatika zapnuta a zařízení připraven pro nový příkaz od PC", 3));
                cbStatut.Items.Add(new Item("chybně zadané parametry, musí se sepnou Reset pro akceptaci chyby", 4));
                cbStatut.Items.Add(new Item("chyba", 5));
                cbStatut.SelectedIndex = 3;

                cbInfo.Items.Add(new Item("Automatický provoz je vypnutý", 0));
                cbInfo.Items.Add(new Item("Probíhá základní nastavení", 1));
                cbInfo.Items.Add(new Item("Připraven, čeká na příkaz od PC", 2));
                cbInfo.Items.Add(new Item("Kontrola příkazu od PC", 3));
                cbInfo.Items.Add(new Item("Zakládání dílu", 4));
                cbInfo.Items.Add(new Item("Přesun k zakládání", 5));
                cbInfo.Items.Add(new Item("Přesun ke kameře", 6));
                cbInfo.Items.Add(new Item("Kontrola orientace", 7));
                cbInfo.Items.Add(new Item("Přesun do zmetkovníku", 8));
                cbInfo.Items.Add(new Item("Přesun k tiskárně", 9));
                cbInfo.Items.Add(new Item("Tisk", 10));
                cbInfo.Items.Add(new Item("Přesun k razníku", 11));
                cbInfo.Items.Add(new Item("Ražení", 12));
                cbInfo.Items.Add(new Item("Přesun do zásobníku OK dílů", 13));
                cbInfo.Items.Add(new Item("HOTOVO, přesun do základní polohy", 14));
                cbInfo.Items.Add(new Item("Řízení vypnuto", 15));
                cbInfo.SelectedIndex = 14;
                
                cbError.Items.Add(new Item("Bez chyby", 0));
                cbError.Items.Add(new Item("Procesorová jednotka zastavena", 1));
                cbError.Items.Add(new Item("Řízení vypnuto", 8));
                cbError.Items.Add(new Item("Ochrany přemostěny", 9));
                cbError.Items.Add(new Item("ESTOP zmáčknut", 10));
                cbError.Items.Add(new Item("Kryt zařízení otevřen", 11));
                cbError.Items.Add(new Item("Nízký tlak", 12));
                cbError.Items.Add(new Item("Nedojel válec", 15));
                cbError.Items.Add(new Item("Chybně zadaný typ", 22));
                cbError.Items.Add(new Item("Chybně zadané jméno", 23));
                cbError.Items.Add(new Item("Chybně zadané os. číslo", 24));
                cbError.Items.Add(new Item("Chyba v zakládání", 25));
                cbError.Items.Add(new Item("Vstupní zásobník dílů prázdný", 27));
                cbError.Items.Add(new Item("Chybně zadaný čárový kód", 31));
                cbError.Items.Add(new Item("Chybně zadaný ražený kód", 32));
                cbError.Items.Add(new Item("Chyba v komunikaci s tiskárnou", 33));
                cbError.Items.Add(new Item("Chyba v komunikaci s razníkem", 34));
                cbError.Items.Add(new Item("Zakládání nastavení nedokončeno", 35));
                cbError.Items.Add(new Item("Chyba portálu", 36));
                cbError.Items.Add(new Item("Vložte cartridge CART do tiskárny", 37));
                cbError.SelectedIndex = 0;

                cbFinishOK.Items.Add(new Item("Ok", 1));
                cbFinishOK.Items.Add(new Item("False", 0));
                cbFinishOK.SelectedIndex = 0;
#endif




                Globalni.Nastroje.LogMessage("Start", false, "Information", formRaz);
                if (!Init())
                {
                    // pokud neprojde Init() nema smysel se ptat dal ..
                    //Vlastnosti.popisStavuRaznice popisStavuRaznice;
                    //popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    //popisStavuRaznice = DejPopisStavu();

                    MessageBox.Show("Chyba při inicializování komunikace s PLC", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Globalni.Nastroje.LogMessage("Chyba při inicializování komunikace s PLC", false, "Error", formRaz);
                    //this.Close();
                    chkReady.Checked = false;
                }
                else
                {
                    Vlastnosti.popisStavuRaznice popisStavuRaznice;
                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu();
                    if ((popisStavuRaznice.nStatusId == 3)) //zařízení zapnuto
                        this.chkReady.Checked = true;
                    else
                    {
                        this.chkReady.Checked = false;
                        MessageBox.Show("Load Init(): " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Globalni.Nastroje.LogMessage("Load Init(): " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                    }

                    txtRazitDoz.Text = "";
                    txtRazitDoz.PromptChar = ' ';
                    txtRazitDoz.Mask = "000";

                    txtRazitOdDoz.Text = "";
                    txtRazitOdDoz.PromptChar = ' ';
                    //txtRazitOdDoz.Mask = "00 000 000";

                    timer1.Enabled = true;
                    //this.chkReady.Checked = true;
                }
            }
            catch
            {
                MessageBox.Show("Nebyla nalezena knihovna RazniceV2.dll", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("Nebyla nalezena knihovna RazniceV2.dll", false, "Error", formRaz);
                this.Close();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globalni.Nastroje.LogMessage("Stop", false, "Information", formRaz);
            timer1.Enabled = false;

            try
            {
                //Disconnect();
            }
            catch (Exception ex)
            {
                string chyba = "Source:" + ex.Source.ToString() +
                                                ", Message:" + ex.Message.ToString() +
                                                ", Stack:" + ex.StackTrace.ToString() +
                                                ", TargetSite:" + ex.TargetSite.ToString() +
                                                ", Data:" + ex.Data.ToString();
                Globalni.Nastroje.LogMessage("Raznice: " + chyba, false, "Error", formRaz);
            }
           
        } 

#endregion

#region Ovladaci_prvky

        /// <summary>
        /// Nastaveni Ready a zacatku nejakeho razeni
        /// </summary>
        /// <param name="ready"></param>
        private void EnablingReady(bool ready)
        {
            //chkReady.Checked = ready;
            txtSarze.Enabled = ready;
            cbTypFilmu.Enabled = ready;

            // zalozka Z tabulky            
            cmdOtevritPlan.Enabled = ready;
            cmdOznacitVseVyrazit.Enabled = ready;
            cmdOdeznacitVseVyrazit.Enabled = ready;
            dataGridView1.Enabled = ready;
            dataGridView1.UseWaitCursor = !ready;

            cmdVyrazit.Enabled = ready; // vyrazit

            // zalozka Postupna
            btnStart.Enabled = ready;   // vyrazit dozimetr

            // zalozka Ze souboru
            btnStartFromFile.Enabled = ready; // vyrazit ze souboru
            btnLoadFile.Enabled = ready;


        }

        /// <summary>
        /// Nastaveni Done - hotovo 
        /// </summary>
        /// <param name="ready"></param>
        private void EnablingDone(bool ready)
        {
            chkDone.Checked = ready;
            //btnLoadFile.Enabled = ready;

        }

  

        /// <summary>
        /// Tab "postupna", Kontrola, zda je vse vyplneno
        /// </summary>
        /// <returns></returns>
        private bool Kontrola()
        {
            string txt = txtText.Text.Trim();

#region kontrola

            if (txtSarze.Text == String.Empty)
            {
                MessageBox.Show("Šarže filmu není zadána", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSarze.Focus();
                return false;

            }

            if (txtTyp.Text == String.Empty)
            {
                MessageBox.Show("Typ filmu není zadán", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return false;

            }
            if (txtTyp.Text != "1" && txtTyp.Text != "2" && txtTyp.Text != "3")
            {
                MessageBox.Show("Typ filmu není zadán v intervalu 1 - 3", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return false;

            }
            int nTyp = 0;
            if (!int.TryParse(txtTyp.Text, out nTyp))
            {
                MessageBox.Show("Typ filmu není zadán korektně v intervalu 1 - 3", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return false;

            }

            // kontrola vyplneni

            if (txt == String.Empty)
            {
                MessageBox.Show("Číslo dozimetru (ražené číslo) musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtText.Focus();
                return false;
            }

            if (((txt.Length != 8) && (nTyp == 2)) || ((txt.Length != 8) && (nTyp == 3)))
            {
                MessageBox.Show("Číslo dozimetru (ražené číslo) musí být 8 znaků pro typ 2, 3.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtText.Focus();
                return false;
            }
            if ((txt.Length != 6) && (nTyp == 1))
            {
                MessageBox.Show("Číslo dozimetru (ražené číslo) musí být 6 znaků pro typ 1.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtText.Focus();
                return false;
            }

            if (txtObdobi.Text == String.Empty)
            {
                MessageBox.Show("Číslo období musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtObdobi.Focus();
                return false;
            }
            if (txtMesic.Text == String.Empty)
            {
                MessageBox.Show("Číslo měsíce musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMesic.Focus();
                return false;
            }
            if (txtRok.Text == String.Empty)
            {
                MessageBox.Show("Číslo roku musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtRok.Focus();
                return false;
            }
            if (txtPodnik.Text == String.Empty)
            {
                MessageBox.Show("Číslo podniku musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPodnik.Focus();
                return false;
            }
            if (txtOddeleni.Text == String.Empty)
            {
                MessageBox.Show("Číslo oddělení podniku musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtOddeleni.Focus();
                return false;
            }
            if (txtDozimetr.Text == String.Empty)
            {
                MessageBox.Show("Číslo dozimetru musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtDozimetr.Focus();
                return false;
            }

            //          if (txtJmeno.Text == String.Empty)
            //          {
            //              MessageBox.Show("Jméno musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //              return;
            //          }

            int numero = 0;
            // kotrola na delku
            if ((txtObdobi.Text.Length != 1) || !(int.TryParse(txtObdobi.Text, out numero)))
            {
                MessageBox.Show("Číslo období musí být vyplněno jednou číslicí.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtObdobi.Focus();
                return false;
            }
            if ((txtMesic.Text.Length != 2) || !(int.TryParse(txtMesic.Text, out numero)))
            {
                MessageBox.Show("Číslo měsíce musí být vyplněno dvěma číslicemi.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMesic.Focus();
                return false;
            }
            if ((txtRok.Text.Length != 2) || !(int.TryParse(txtRok.Text, out numero)))
            {
                MessageBox.Show("Číslo roku musí být vyplněno dvěma číslicemi.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtRok.Focus();
                return false;
            }
            if ((txtPodnik.Text.Length != 3) || !(int.TryParse(txtPodnik.Text, out numero)))
            {
                MessageBox.Show("Číslo podniku musí být vyplněno třemi číslicemi.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPodnik.Focus();
                return false;
            }
            if ((txtOddeleni.Text.Length != 1) || !(int.TryParse(txtOddeleni.Text, out numero)))
            {
                MessageBox.Show("Číslo oddělení podniku musí být vyplněno jednou číslicí.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtOddeleni.Focus();
                return false;
            }
            if ((txtDozimetr.Text.Length != 3) || !(int.TryParse(txtDozimetr.Text, out numero)))
            {
                MessageBox.Show("Číslo dozimetru musí být vyplněno třemi číslicemi.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtDozimetr.Focus();
                return false;
            }


            if ((txtJmeno.Text == String.Empty))
            {
                MessageBox.Show("Jméno musí být uvedeno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtJmeno.Focus();
                return false;
            }
            if ((txtJmeno.Text.Length > 14))
            {
                MessageBox.Show("Jméno nesmí být delší než 14 znaků.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtJmeno.Focus();
                return false;
            }

#endregion

            return true;
        }

        /// <summary>
        /// Tab2 "postupna", Vyrazit dozimetr 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel.Text = "";
            Vlastnosti.popisStavuRaznice popisStavuRaznice;
            if (!Init())
            {
                MessageBox.Show("chyba Init()", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("chyba Init()", false, "Error", formRaz);
                chkReady.Checked = false;
                return;
            }
            chkReady.Checked = true;

            popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
            popisStavuRaznice = DejPopisStavu();
            if ((popisStavuRaznice.nStatusId != 3)) //chyba, řízení vypnuto
            {
                MessageBox.Show("nStatusId: " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("nStatusId: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                return;
            }

           

            //string txt = InsertSpace(txtText.Text);
            string txt = txtText.Text.Trim();

            // Tisk
            string cislo_ean = "";
            string popisek_stitku = "";

            // kontrola vyplneni udaju na tab "postupna"
            if (!Kontrola())
                return;

            // 1 06 16 130 2 203

            // 1A_06_130/2_203
            popisek_stitku = txtObdobi.Text.Trim() + txtSarze.Text.Trim() + '_' + // 1A
                             txtMesic.Text.Trim() + '_' + //  06
                             txtPodnik.Text.Trim() + "/" + txtOddeleni.Text.Trim() + '_' + // 130/2
                             txtDozimetr.Text.Trim();   // 203
            // Vachata
            popisek_stitku = popisek_stitku + " " + txtJmeno.Text.Trim();

            // 106151302203
            cislo_ean = txtObdobi.Text.Trim() + // 1
                             txtMesic.Text.Trim() + // 06  
                             txtRok.Text.Trim() + // 15
                             txtPodnik.Text.Trim() + txtOddeleni.Text.Trim() + // 1302
                             txtDozimetr.Text.Trim();   // 203

            // nastavim si, protoze v NaRazitDozV2 se dle toho ridi, zda vubec vyrazit a tady to chci bezpecne
            bool checkedRazeni = chkRazitDozimetry.Checked;
            chkRazitDozimetry.Checked = true;

            bool vysledek = NaRazitDozV2(txt, popisek_stitku,  cislo_ean, txtTyp.Text.ToString());
            chkRazitDozimetry.Checked = checkedRazeni;

            return;
            ///////////////////////////////////////////
            #region primarne definovany postup razeni v Trebici 
            //// tisk popisku z tab. Postupna 
            //// nastavi se vse potrebne
            //int kolecko = 1;
            //string nTyp = txtTyp.Text;
            //while (kolecko <= 3)
            //{
            //    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
            //    popisStavuRaznice = DejPopisStavu();
            //    if ((popisStavuRaznice.nStatusId != 3)) //neni chyba, neni řízení vypnuto
            //    {
            //        MessageBox.Show("nStatusId: " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        Globalni.Nastroje.LogMessage("nStatusId: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);

            //        Cekej(2);
            //        kolecko++;
            //        continue;
            //    }

            //    if (!SetTiskV2(nTyp.ToString() /*2*/, txt, popisek_stitku, cislo_ean, false, true))
            //    {
            //        popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
            //        popisStavuRaznice = DejPopisStavu();

            //        MessageBox.Show("SetTiskV2: " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        Globalni.Nastroje.LogMessage("SetTiskV2: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
            //    }

            //    if (!Start())
            //    {
            //        MessageBox.Show("chyba Start()", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        Globalni.Nastroje.LogMessage("chyba Start()", false, "Error", formRaz);
            //    }

            //    Cekej(1);

            //    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
            //    popisStavuRaznice = DejPopisStavu();

            //    if (popisStavuRaznice.nStatusId == 4)
            //    {
            //        MessageBox.Show("po SetTiskV2 nStatusId == 4: " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        Globalni.Nastroje.LogMessage("po SetTiskV2 nStatusId == 4: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);

            //        if (!Reset())
            //        {
            //            MessageBox.Show("chyba pri Reset()", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            Globalni.Nastroje.LogMessage("chyba pri Reset()", false, "Error", formRaz);
            //        }

            //        Cekej(1);
            //        kolecko++;
            //        continue;
            //    }
            //    else
            //    {
            //        // koncim cyklus sem ok a jdu do finise
            //        break;
            //    }
            //}

            //int koleckoFinish = 0;
            //while (koleckoFinish <= 3)
            //{
            //    Cekej(2);
            //    bool lOk = false;
            //    if (!ReadFinishOK(ref lOk))
            //    {
            //        MessageBox.Show("chyba pri ReadFinishOK()", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //        Globalni.Nastroje.LogMessage("chyba pri ReadFinishOK()", false, "Error", formRaz);
            //    }

            //    if (lOk == false)
            //    {

            //        popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
            //        popisStavuRaznice = DejPopisStavu();

            //        // pokud je bez chyby, znovu
            //        if (popisStavuRaznice.nStatusId == 5)
            //        {
            //            MessageBox.Show("po !lOk popisStavuRaznice.nStatusId == 5: " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            Globalni.Nastroje.LogMessage("po !lOk popisStavuRaznice.nStatusId == 5: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);

            //            // ctu error, ten ale mam uz nacteny
            //            MessageBox.Show("po !lOk popisStavuRaznice.nErroId: " + popisStavuRaznice.nErrorId.ToString() + " -" + popisStavuRaznice.nErrorText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            //            Globalni.Nastroje.LogMessage("po !lOk popisStavuRaznice.nErroId: " + popisStavuRaznice.nErrorId.ToString() + " -" + popisStavuRaznice.nErrorText.ToString(), false, "Error", formRaz);

            //            // KONCIM
            //            break;
            //        }
            //        else
            //        {
            //            Globalni.Nastroje.LogMessage("po !lOk popisStavuRaznice: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
            //            koleckoFinish++;
            //            continue;
            //        }
            //    }
            //    else
            //    {
            //        // je finis OK, mam narazeno a vytisklo, jdu ven
            //        Globalni.Nastroje.LogMessage("po lOk popisStavuRaznice: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
            //        break;
            //    }


            //}
            #endregion
        }




        /// <summary>
        /// z tab3 "ze souboru" Vyrazit ze souboru
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStartFromFile_Click(object sender, EventArgs e)
        {
            // razeni dle textaku
            toolStripStatusLabel.Text = "";
            DozFile = true;
            StartNV2();
        }

        /// <summary>
        /// z tab3 "ze souboru" Otevrit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            //OpenDialog.InitialDirectory = "./.";
            //OpenDialog.FileName = "./";
            lblCelkem.Text = "0";

            OpenDialog.Filter = "Textové soubory (*.txt)|*.txt";
            if (OpenDialog.ShowDialog() == DialogResult.OK)
            {
                txtFile.Text = OpenDialog.FileName;
            }
            DozStr = LoadFile(txtFile.Text);
            if (!(DozStr == null))
            {                
                RozeberDozStr(0);
                lblCelkem.Text = DozStr.Count().ToString();

                /*
                lblDozNum.Text = DozStr[0];
                 */
            }
            else MessageBox.Show("Soubor se nepodařilo načíst", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


        private void btnStop_Click(object sender, EventArgs e)
        {
#if DLL
            //if (btnStop.Text == "STOP") { Stop(); }
            //else { Run(); }

            if (btnStop.Text == "STOP") 
            { 
                //Stop();
                vProcesuRazeni = false;
                EnablingReady(true);
            }
            else 
            { 
                //Run(); 
            }
#else
            //simulace fci z Raznice.dll
            //timer2.Enabled = false;
            //vProcesuRazeni = false;
            //EnablingReady(true);

            if (btnStop.Text == "STOP") 
            { 
                //Stop();
                vProcesuRazeni = false;
                EnablingReady(true);
            }
            else 
            { 
                //Run(); 
            }

#endif
        }

        /// <summary>
        /// Prestaveni tl. STOP/RUN
        /// </summary>
        /// <param name="stop"></param>
        private void STPbtn(bool stop)
        {
            if (stop)
            {
                btnStop.Text = "STOP";
                btnStop.BackColor = Color.Red;
            }
            else
            {
                btnStop.Text = "RUN";
                btnStop.BackColor = Color.Green;
            }
        }
        
        /// <summary>
        /// Na strip pasu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReconnect_ButtonClick(object sender, EventArgs e)
        {
            btnStop.Enabled = true;
            timer1.Enabled = true;
            btnReconnect.Visible = false;
        }

        #endregion

        #region Timery

        /// <summary>
        /// Sleduje stav raznice
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            bool ready = false, done = false;
            int Err = 0, Mark = 0;
            bool ok = false;
            Vlastnosti.popisStavuRaznice popisStavuRaznice = null;

            popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
            popisStavuRaznice = DejPopisStavu(nelogovat: true);

            // v pripade, ze nezjistim stav
            #region nelze zjistit stav raznice
            if (popisStavuRaznice.nStatusId == -1 || popisStavuRaznice.nErrorId == -1 || popisStavuRaznice.nInfoId == -1)
            {
                // neznamy stav
                ok = false;
            }
            else
                ok = true;
            // neni-li mozne zjistit stav raznice
            if (!ok)
            {
                STPbtn(true);
                lblStatus.Text = "Chyba komunikace, nelze zjistit stav";
                timer1.Enabled = false;
                btnStop.Enabled = false;
                btnReconnect.Visible = true;
                EnablingDone(false); /*false*/
                EnablingReady(false);
                Globalni.Nastroje.LogMessage("timer1_Tick, lblStatus.Text " + lblStatus.Text.ToString(), false, "Error", formRaz);

                return;
            }
            #endregion

            #region OLD vyhodnoceni orazeni posledniho dozimetru
            //else
            //{
            //    done = (popisStavuRaznice.nStatusId == 3);
            //    if (timer2.Enabled) 
            //    { 
            //        EnablingDone(false); 
            //    } 
            //    else 
            //        EnablingDone(done);

            //    Err = (popisStavuRaznice.nErrorId);
            //    if (Err > 0)
            //    {
            //        STPbtn(false);
            //        lblStatus.Text = "Error: " + popisStavuRaznice.nErrorText;

            //        //switch (Mark)
            //        //{
            //        //    case 0:
            //        //        lblMark.Text = "- Poslední dozimetr neoražen";                            
            //        //        break;
            //        //    case 1:
            //        //        lblMark.Text = "- Poslední dozimetr oražen";
            //        //        break;
            //        //    case 2:
            //        //        lblMark.Text = "- Nelze zjistit, je li poslední dozimetr oražen správně";
            //        //        break;
            //        //}
            //        lblMark.Text = popisStavuRaznice.nErrorText;
            //        Globalni.Nastroje.LogMessage("timer1_Tick, lblStatus.Text " + lblStatus.Text.ToString() + ", lblMark.Text: " + lblMark.Text.ToString(), false, "Error", formRaz);
            //    }
            //    else
            //    {
            //        STPbtn(true);
            //        lblStatus.Text = "";
            //        lblMark.Text = "";
            //    }
            //}
            #endregion

            lblStatus.Text = popisStavuRaznice.stavText.ToString();

            // umoznim pokracovat v cyklu razeni, mel bych mit nastaveno kde jsem bud v dbf nebo v poctu dozimetru DozCount
            if (popisStavuRaznice.nStatusId == 5)
            {
                // zpristupnim STOP
                STPbtn(true);
            }

            ready = (popisStavuRaznice.nStatusId == 3);

            

            if (vProcesuRazeni) /*(timer2.Enabled)*/
            { 
                EnablingReady(false); 
            } else 
                EnablingReady(ready);

              
        }
  
#endregion

        private void txtSarze_LostFocus(object sender, EventArgs e)
        {
            txtSarze.Text = txtSarze.Text.ToUpper();
        }

#region Charset

        public static byte[] StringToByteArray(string hex)
        {
             return Enumerable.Range(0, hex.Length)
                     .Where(x => x % 2 == 0)
                     .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                     .ToArray();
        }

        private string DecodeISO8859_1(string str)
        {
            var text = Regex.Replace(str, "=([0-9A-F][0-9A-F])", delegate(Match matchChar)
            {
                return Encoding.GetEncoding("iso-8859-1").GetString(StringToByteArray(matchChar.Groups[1].Value));
            });
            return text;
        }

        private string DecodeISO8859_2(string str)
        {
            var text = Regex.Replace(str, "=([A-F][0-9A-F])|=([0-9][0-9A-F])", delegate(Match matchChar)
            {
                var hex = Encoding.GetEncoding("iso-8859-2").GetString(StringToByteArray(matchChar.Groups[1].Value));
                if (hex == "") hex = Encoding.GetEncoding("iso-8859-2").GetString(StringToByteArray(matchChar.Groups[2].Value));
                return hex;
            });
            return text;
        }

        private string DecodeWindows1250(string str)
        {
            var text = Regex.Replace(str, "=([0-9A-F][0-9A-F])", delegate(Match matchChar)
            {
                return Encoding.GetEncoding("windows-1250").GetString(StringToByteArray(matchChar.Groups[1].Value));
            });
            return text;
        }

        private string DecodeUTF8(string str)
        {
            var text = Regex.Replace(str, "=([C][0-9A-F])=([0-9A-F][0-9A-F])|=([C][0-9A-F])==([0-9A-F][0-9A-F])|=([0-9A-F][0-9A-F])",
              delegate(Match matchChar)
              {
                  var hex = Encoding.UTF8.GetString(StringToByteArray(matchChar.Groups[1].Value + matchChar.Groups[2].Value));
                  if (hex == "") hex = Encoding.UTF8.GetString(StringToByteArray(matchChar.Groups[3].Value + matchChar.Groups[4].Value));
                  else if (hex == "") hex = Encoding.UTF8.GetString(StringToByteArray(matchChar.Groups[5].Value));
                  return hex;
              });
            return text;
        }

        private string Decodecharset(string str)
        {
            //charset Base64
            str = Regex.Replace(str, @"=\?[uUtTfF]+-8\?[bB]\?([a-zA-Z0-9]+={0,2})\?=",
                       delegate(Match match)
                       {
                           var bytes = Convert.FromBase64String(match.Groups[1].Value);
                           return Encoding.UTF8.GetString(bytes);
                       });

            //charset iso-8859-1
            str = Regex.Replace(str, @"=\?[iIsSoO]+-8859-1\?[qQ]\?(.+)\?=",
                       delegate(Match match)
                       {
                           return DecodeISO8859_1(match.Groups[1].Value);
                       });

            //charset iso-8859-2
            str = Regex.Replace(str, @"=\?[iIsSoO]+-8859-2\?[qQ]\?(.+)\?=",
                       delegate(Match match)
                       {
                           return DecodeISO8859_1(match.Groups[1].Value);
                       });

            //charset windows-1250
            str = Regex.Replace(str, @"=\?[wWiInNdDoOwWsS]+-1250\?[qQ]\?(.+)\?=",
                       delegate(Match match)
                       {
                           return DecodeWindows1250(match.Groups[1].Value);
                       });

            //charset utf8
            str = Regex.Replace(str, @"=\?[uUtTfF]+-8\?[qQ]\?(.+)\?=",
                       delegate(Match match)
                       {

                           return DecodeUTF8(match.Groups[1].Value);
                       });
            return str;
        }

#endregion

#region funkce z tab postupna a volny
        /// <summary>
        /// Nastaveni textu pro razbu V2 a potisk, neni pro plan 
        /// </summary>
        /// <param name="txt_DozNum"></param>
        /// <param name="popisek_stitek"></param>
        /// <param name="cislo_ean"></param>
        /// <param name="hlasitChybu"></param>
        /// <param name="VolnyTisk"></param>
        /// <returns></returns>
        private bool SetTiskV2(string typeDoz /*1,2,3*/, string txt_DozNum /*11001003*/, string popisek_stitek /*1A_06_130/2_203 Michlova*/, string cislo_ean /*106151302203*/, bool hlasitChybu, bool VolnyTisk)
        {
            bool jakTisk = false;
            string nameZdroj = "";
            string numZdroj = "";
            string namePrint = "";
            string personalNoPrint = "";
            bool jakSendText = false;
            Vlastnosti.popisStavuRaznice popisStavuRaznice = null;

            //1A Michlova
            //050190002

            // od 05.04.2016 obsahuje i cislo oddeleni: 0PPPDDD--> 0PPPODDD
            // Vejsada
            // 0PPPDDD

            try
            {
                nameZdroj = popisek_stitek.Trim(); //1A_06_130/2_203 Michlova
                numZdroj = cislo_ean.Trim();       //106151302203
                var rows = nameZdroj.Split(' ');
                if (rows != null)
                {
                    personalNoPrint = rows[0];
                    namePrint = rows[1];
                }

                if (!VolnyTisk)
                {
                    if ((numZdroj.Length != 12) && (hlasitChybu))
                        MessageBox.Show("Číslo pro konstrukci EAN kódu: '" + numZdroj.ToString() + "' musí být dlouhé 12 znaků.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (numZdroj.Length != 12)
                    {
                        Globalni.Nastroje.LogMessage("SetTiskV2, Číslo pro konstrukci EAN kódu musí být dlouhé 12 znaků, numZdroj:" + numZdroj.ToString(), false, "Error", formRaz);
                        return false;
                    }

                    //1A_06_130/2_203
                    if ((nameZdroj.Length < 15) && (hlasitChybu))
                        MessageBox.Show("Text štítku dozimetru '" + numZdroj.ToString() + "' musí být minimálně 15 znaků dlouhý.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // pokud je mensi, doplnim na 15 pozice - je mozne tisknout jen cislo dozimetru beze jmena
                    if (nameZdroj.Length < 15)
                    {
                        Globalni.Nastroje.LogMessage("SetTiskV2, Text štítku dozimetru musí být minimálně 15 naků dlouhý., numZdroj:" + nameZdroj.ToString(), false, "Warning", formRaz);
                        nameZdroj = nameZdroj.PadLeft(15, '0');
                    }
                }
                else
                {
                    // doplnim zleva na 12 znaku pro EAN13
                    if (numZdroj.Length != 12)
                    {
                        numZdroj = numZdroj.PadLeft(12, '0');
                    }
                }

                if (nameZdroj.Length > 30)
                    nameZdroj = nameZdroj.Substring(0, 30);


                Globalni.Nastroje.LogMessage("SetTiskV2 SendTextName:" + namePrint.ToString() + ", SendTextPersonalNo: " + personalNoPrint.ToString() + ", SendTextBarCode: " + numZdroj.ToString() + ", SendTextRazNo: " + txt_DozNum.ToString(), false, "Information", formRaz);
                //jakTisk = PrintEAN13(numZdroj, numZdroj.Length, nameZdroj, nameZdroj.Length);

                //if (SendType(typeDoz))
                //    if (SendTextName(namePrint, namePrint.Length))
                //        if (SendTextPersonalNo(personalNoPrint, personalNoPrint.Length))
                //            if (SendTextBarCode(numZdroj, numZdroj.Length))
                //                if (SendTextRazNo(txt_DozNum, txt_DozNum.Length))
                //                    jakTisk = true;

#region metody SendText
                jakSendText = true;
                //char cType = char.Parse(typeDoz.ToString());
                //if (!SendType(cType)
                if (!SendType(typeDoz.ToString() /*cType*/))
                {
                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu();
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendType(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                    jakSendText = false;
                }
                if (!SendTextName(namePrint, namePrint.Length))
                {
                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu();
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextName(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                    jakSendText = false;
                }
                if (!SendTextPersonalNo(personalNoPrint, personalNoPrint.Length))
                {
                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu();
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextPersonalNo(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                    jakSendText = false;
                }
                if (!SendTextBarCode(numZdroj, numZdroj.Length))
                {
                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu();
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextBarCode(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                    jakSendText = false;
                }
                if (!SendTextRazNo(txt_DozNum, txt_DozNum.Length))
                {
                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu();
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextRazNo(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                    jakSendText = false;
                }
#endregion

                if (jakSendText == true)
                {
                    jakTisk = true;
                    lblStatus.Text = "Tisk ok";
                    //toolStripStatusLabel.Text = "Tisk EAN13 ok";
                }
                else
                {
                    jakTisk = false;
                    lblStatus.Text = "Chyba SetTiskV2";
                    toolStripStatusLabel.Text = "Chyba SetTiskV2";
                    if (hlasitChybu)
                        MessageBox.Show("Chyba SetTiskV2", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                Globalni.Nastroje.LogMessage("SetTiskV2 " + lblStatus.Text, false, "Information", formRaz);
                MessageBox.Show("SetTiskV2 SendTextName:" + namePrint.ToString() + ", SendTextPersonalNo: " + personalNoPrint.ToString() + ", SendTextBarCode: " + numZdroj.ToString() + ", SendTextRazNo: " + txt_DozNum.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            catch (Exception ex)
            {
                Globalni.Nastroje.LogMessage("Chyba SetTiskV2(): " + ex.Message.ToString(), false, "Error", formRaz);
            }
            return jakTisk;
        }


#endregion

#region funkce z tab plan
        private int indexOf(DataGridView dgv, string name) 
        {
            int index = 0;
            try
            {
                index = dgv.Columns[name].Index;
            }
            catch
            {
                Globalni.Nastroje.LogMessage("Nenalezen column name: " + name + " pro dgv: " + dgv.Name + " ?", false, "Error", formRaz);
                MessageBox.Show("Nenalezen column name: " + name + " pro dgv: " + dgv.Name + " ?", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                index = 0;
            }


            return index; 
        } 
             
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {            
            try
            {
                if (e.ColumnIndex == indexOf(dataGridView1, "Vyrazit")) // chkbox Vyrazit                    
                {
                    int rowindex = dataGridView1.CurrentCell.RowIndex;
                    int columnindex = dataGridView1.CurrentCell.ColumnIndex;

                    //string a = dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();

                    int id_cispod = Int32.Parse(dataGridView1[indexOf(dataGridView1, "Id_Cispod"), e.RowIndex].Value.ToString());
                    int vyrazit = Int32.Parse(dataGridView1[indexOf(dataGridView1, "Vyrazit"), e.RowIndex].Value.ToString());

                    if (vyrazit == 1)
                        vyrazit = 0;
                    else
                        vyrazit = 1;
                    //dataGridView1.DataSource = "";

                    dataGridView1[indexOf(dataGridView1, "Vyrazit"), e.RowIndex].Value = vyrazit;
                    DataTable ResultSet = UpdateGRPDataVyrazit(id_cispod, vyrazit);
                    //NastavDataGrid(dataGridView1);
                    //dataGridView1.DataSource = ResultSet;

                    //dataGridView1.Rows[rowindex].Selected = true;
                    //dataGridView1.CurrentCell = dataGridView1.Rows[rowindex].Cells[0];
 
                }

                if (e.ColumnIndex == indexOf(dataGridView1,"Nacist")) // button Nacist    
                                
                {
                    lblPodnikuProVyrazeni.Text = "0";

                    int rowindex = dataGridView1.CurrentCell.RowIndex;
                    int columnindex = dataGridView1.CurrentCell.ColumnIndex;
                    //string a = dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();

                    int id_cispod = Int32.Parse(dataGridView1[indexOf(dataGridView1, "Id_Cispod"), e.RowIndex].Value.ToString());
                    int zpracovano = Int32.Parse(dataGridView1[indexOf(dataGridView1, "Zpracovano"), e.RowIndex].Value.ToString());
                    int vyrazit = Int32.Parse(dataGridView1[indexOf(dataGridView1, "Vyrazit"), e.RowIndex].Value.ToString());

                    if (zpracovano == 1)
                    {
                        //MessageBox.Show("Uz je zpracovano");
                        string cpd = dataGridView1[indexOf(dataGridView1, "cpd"), e.RowIndex].Value.ToString();
                        string cod = dataGridView1[indexOf(dataGridView1, "cod"), e.RowIndex].Value.ToString();  

                        DialogResult result = MessageBox.Show("Pro podnik "+cpd+"/"+cod+" je již vše naraženo. \r\nPokračovat?", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        switch (result)
                        {
                            case DialogResult.Yes:
                                {
                                    // odeznacit zpracovano, aby doslo k vyberu Odebrat označení?
                                    //DataTable ResultSetUpdate = UpdateGRPData(id_cispod, zpracovano: 0);
                                    break;
                                }
                            case DialogResult.No:
                                {
                                    return;
                                    //break;
                                }
                        }
                    }
                    // v pripade, ze neni oznaceno k vyrazeni
                    if (vyrazit == 0)
                    {
                        DataTable ResultSetUpdateVyrazit = UpdateGRPDataVyrazit(id_cispod, vyrazit: 1);
                        dataGridView1[indexOf(dataGridView1, "Vyrazit"), e.RowIndex].Value = 1;
                    }

                    ////zapsat pro vsechny zaznamy "Vyrazit" do tabulky, jak je v gridu
                    //foreach (DataGridViewRow row in dataGridView1.Rows)
                    //{
                    //    decimal hodnotaVyrazit = 0;
                    //    hodnotaVyrazit = (decimal)row.Cells[indexOf(dataGridView1, "Vyrazit")].Value;
                        
                    //    if (hodnotaVyrazit == 1)
                    //        UpdateGRPDataVyrazit(id_cispod, vyrazit: 1);

                    //}


                    dataGridView1.DataSource = "";


                    //DataTable ResultSet = UpdateGRPData(id_cispod);                    
                    DataTable ResultSet = GetGRPData();
                    NastavDataGrid(dataGridView1);
                    dataGridView1.DataSource = ResultSet;

                    dataGridView1.Rows[rowindex].Selected = true;
                    dataGridView1.CurrentCell = dataGridView1.Rows[rowindex].Cells[0];


                    // nactu pro vsechny podniky, oznacene Vyrazit
                    dataGridView2.DataSource = "";
                    DataTable ResultSetCDZ = GetDOZData();
                    NastavDataGrid(dataGridView2);
                    dataGridView2.DataSource = ResultSetCDZ;

                    dataGridView2.Rows[0].Selected = true;
                    dataGridView2.CurrentCell = dataGridView1.Rows[0].Cells[0];

                    //dataGridView2_CellContentClick(sender, e);
                    string Tisk_radek_1 = (dataGridView2[indexOf(dataGridView2, "Tisk_radek_1"), 0]).Value.ToString();
                    string Tisk_radek_2 = (dataGridView2[indexOf(dataGridView2, "Tisk_radek_2"), 0]).Value.ToString();
                    string Tisk_prijmeni = (dataGridView2[indexOf(dataGridView2, "PRIJMENI"), 0]).Value.ToString();
                    // 05.04.2016 doplneno tisk COD do eanu, zmena eanu z EAN8 na EAN13
                    string Tisk_cod =      (dataGridView2[indexOf(dataGridView2, "Oddeleni"), 0]).Value.ToString();
                    string Tisk_slob = (dataGridView2[indexOf(dataGridView2, "SLOB"), 0]).Value.ToString();
                    string Tisk_rok = (dataGridView2[indexOf(dataGridView2, "RP_ROK"), 0]).Value.ToString();
                    string Tisk_mesic = (dataGridView2[indexOf(dataGridView2, "RP_MESIC"), 0]).Value.ToString();

                    NastavPopisDoz(Tisk_radek_1, Tisk_radek_2, Tisk_prijmeni, Tisk_cod, Tisk_slob, Tisk_rok, Tisk_mesic);

                    List<int> seznamId_Cispod = DejSeznamPodnikuProVyrazeni();
                    lblPodnikuProVyrazeni.Text = seznamId_Cispod.Count().ToString();
                    lblDozimetruVyrazit.Text = GetDozDataZpracovat().ToString();
                }
            }

            catch (Exception ex)
            {
                string chyba = "Source:" + ex.Source.ToString() + ", Message:" + ex.Message.ToString() + ", Data:" + ex.Data.ToString();
                Globalni.Nastroje.LogMessage("Raznice: " + chyba, false, "Error", formRaz);
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
            string Tisk_radek_1 = (dataGridView2[indexOf(dataGridView2, "Tisk_radek_1"), e.RowIndex]).Value.ToString();
            string Tisk_radek_2 = (dataGridView2[indexOf(dataGridView2, "Tisk_radek_2"), e.RowIndex]).Value.ToString();
            string Tisk_prijmeni = (dataGridView2[indexOf(dataGridView2, "PRIJMENI"), e.RowIndex]).Value.ToString();
            // 05.04.2016 doplneno tisk COD do eanu, zmena eanu z EAN8 na EAN13
            string Tisk_cod = (dataGridView2[indexOf(dataGridView2, "Oddeleni"), 0]).Value.ToString();
            string Tisk_slob = (dataGridView2[indexOf(dataGridView2, "SLOB"), 0]).Value.ToString();
            string Tisk_rok = (dataGridView2[indexOf(dataGridView2, "RP_ROK"), 0]).Value.ToString();
            string Tisk_mesic = (dataGridView2[indexOf(dataGridView2, "RP_MESIC"), 0]).Value.ToString();


            NastavPopisDoz(Tisk_radek_1, Tisk_radek_2, Tisk_prijmeni, Tisk_cod, Tisk_slob, Tisk_rok, Tisk_mesic);
        }

        private void cmdOtevritPlan_Click(object sender, EventArgs e)
        {
            //OpenDialog.InitialDirectory = "./.";
            //OpenDialog.FileName = "./";;

            OpenDialog.Filter = "DBF soubory (*GRP*.dbf)|*GRP*.dbf";
            if (OpenDialog.ShowDialog() == DialogResult.OK)
            {
                dbFileName = OpenDialog.FileName;

                //kontrola struktury
                if (!KontrolaGRPData())
                {
                    MessageBox.Show("Soubor " + dbFileName + " se nepodařilo načíst", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    toolStripStatusLabel.Text = "Soubor " + dbFileName + " se nepodařilo načíst";
                    Globalni.Nastroje.LogMessage("cmdOtevritPlan_Click, Soubor " + dbFileName + " se nepodařilo načíst.", false, "Error", formRaz);

                    return;
                }

                DataTable ResultSet = GetGRPData();

                NastavDataGrid(dataGridView1);

                dataGridView1.DataSource = ResultSet;
                dataGridView2.DataSource = "";
                //toolStripStatusLabel.Text = "Soubor " + dbFileName + " byl načten ok.";
                Globalni.Nastroje.LogMessage("cmdOtevritPlan_Click, Soubor " + dbFileName + " byl načten ok.", false, "Information", formRaz);
            }

            else
            {
                MessageBox.Show("Soubor " + dbFileName + " se nepodařilo načíst", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                toolStripStatusLabel.Text = "Soubor " + dbFileName + " se nepodařilo načíst";
                Globalni.Nastroje.LogMessage("cmdOtevritPlan_Click, Soubor " + dbFileName + " se nepodařilo načíst.", false, "Error", formRaz);
            }
        }

        public bool KontrolaGRPData()
        {
            DataTable ResultSet = new DataTable();
            //DataSet ds = new DataSet();

            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";
            OleDbConnection yourConnectionHandler = new OleDbConnection(
                //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            // if including the full dbc (database container) reference, just tack that on
            //      OleDbConnection yourConnectionHandler = new OleDbConnection(
            //          "Provider=VFPOLEDB.1;Data Source=C:\\SomePath\\NameOfYour.dbc;" );


            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                //string mySQL = @"SELECT * FROM 20141015__46B0JSL4X";  // dbf table name
                //string columnName = "";
                try
                {
                    string mySQL = @"SELECT TOP 1 Vyrazit, cpd, cod, kolik, zpracovano, id_cispod FROM " + dbFileName + " ORDER BY cpd, cod";

                    OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                    OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                    DA.Fill(ResultSet);

                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Column 'VYRAZIT' is not found"))
                    {
                        yourConnectionHandler.Close();

                        // ted musim zmenit file, stary prejmenovat na OLD a znej udelat novy puvodniho jmena GRP s rozsirenou strukturou
                        string dbfileOldName = dbFileName.ToUpper().Replace("GRP_", "OLD_");

                        File.Move(dbFileName, dbfileOldName);

                        // Open the connection, and if open successfully, you can try to query it
                        yourConnectionHandler.Open();

                        if (yourConnectionHandler.State == ConnectionState.Open)
                        {
                            string mySQL = @"SELECT 0 as Vyrazit, cpd, cod, kolik, zpracovano, id_cispod FROM " + dbfileOldName + " INTO TABLE " + dbFileName + "";

                            OleDbCommand cmd = new OleDbCommand();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = mySQL;

                            cmd.Connection = yourConnectionHandler;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        // ven
                        throw;
                    }

                }

                yourConnectionHandler.Close();
            }
            return true;
        }

        public DataTable GetGRPData()
        {
            DataTable ResultSet = new DataTable();
            //DataSet ds = new DataSet();

            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";
            OleDbConnection yourConnectionHandler = new OleDbConnection(
                //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            // if including the full dbc (database container) reference, just tack that on
            //      OleDbConnection yourConnectionHandler = new OleDbConnection(
            //          "Provider=VFPOLEDB.1;Data Source=C:\\SomePath\\NameOfYour.dbc;" );


            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                //string mySQL = @"SELECT * FROM 20141015__46B0JSL4X";  // dbf table name
                //string columnName = "";
                try
                {
                    string mySQL = @"SELECT Vyrazit, cpd, cod, kolik, zpracovano, id_cispod FROM " + dbFileName + " ORDER BY cpd, cod";

                    OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                    OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                    DA.Fill(ResultSet);

                    //foreach (DataRow dr in ResultSet.Rows)
                    //{
                    //    //Console.WriteLine(dr.ItemArray[1].ToString());
                    //    columnName = dr.Table.Columns[0].ColumnName;
                    //    break;
                    //}
                }
                catch (Exception ex)
                {
                        throw;

                }

                


                yourConnectionHandler.Close();
            }

            //return ds;
            return ResultSet;
        }

        public DataTable GetDOZData()
        {
            DataTable ResultSet = new DataTable();
            //DataSet ds = new DataSet();

            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";
            OleDbConnection yourConnectionHandler = new OleDbConnection(
                //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            // if including the full dbc (database container) reference, just tack that on
            //      OleDbConnection yourConnectionHandler = new OleDbConnection(
            //          "Provider=VFPOLEDB.1;Data Source=C:\\SomePath\\NameOfYour.dbc;" );


            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                //string mySQL = @"SELECT * FROM 20141015__46B0JSL4X";  // dbf table name
                string fileGRPName = dbFileName;
                string fileName = dbFileName.ToUpper().Replace("GRP_", "");
                //string mySQL = @"SELECT cpd, cod, Cdz, Prijmeni, Tisk_1, Tisk_2, zpracovano, id_seznam, id_cispod, SLOB, RP_ROK, RP_MESIC FROM " + fileName + " where id_cispod = ? ORDER BY cpd, cod, cdz";
                //string mySQL = @"SELECT a.cpd, a.cod, a.Cdz, a.Prijmeni, a.Tisk_1, a.Tisk_2, a.zpracovano, a.id_seznam, a.id_cispod, a.SLOB, a.RP_ROK, a.RP_MESIC FROM " + fileName + " a " +
                string mySQL = @"SELECT a.cpd, a.cod, a.Cdz, a.Prijmeni, a.Tisk_1, a.Tisk_2, a.zpracovano, a.id_seznam, a.id_cispod, a.SLOB, a.RP_ROK, a.RP_MESIC FROM " + fileName + " a " +
                    " JOIN " + fileGRPName +" b ON a.id_cispod = b.id_cispod "+
                    " where (b.Vyrazit = 1 ) ";
                //" where id_cispod = ? ORDER BY cpd, cod, cdz";
                //and b.Zpracovano = 0

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                //OleDbParameter NewParm = new OleDbParameter("id_cispod", id_cispod);
                //NewParm.DbType = DbType.Int32;
                //// (or other data type, such as DbType.String, DbType.DateTime, etc)
                //MyQuery.Parameters.Add(NewParm);

                OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                DA.Fill(ResultSet);
                //DA.Fill(ds);

                /*
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    Console.WriteLine(dr.ItemArray[1].ToString());
                }
                 */
                yourConnectionHandler.Close();
            }

            //return ds;
            return ResultSet;
        }

        public int GetDozDataZpracovat()
        {
            int kolikZpracovat = 0;
            // cyklus pres vsechny oznacene filmy k razeni
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                int hodnotaZpracovano = Int32.Parse((row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value != System.DBNull.Value ? row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value.ToString() : "0"));

                if (hodnotaZpracovano == 0) // kdyz je Zpracovano = 0, tak se jeste nerazil dozimetr
                {
                    kolikZpracovat++;                    
                }
            }
            return kolikZpracovat;
        }

        public DataTable UpdateGRPData(int id_cispod, int zpracovano)
        {
            DataTable ResultSet = new DataTable();
            //DataSet ds = new DataSet();
            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";

            OleDbConnection yourConnectionHandler = new OleDbConnection(
                //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                 @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            // if including the full dbc (database container) reference, just tack that on
            //      OleDbConnection yourConnectionHandler = new OleDbConnection(
            //          "Provider=VFPOLEDB.1;Data Source=C:\\SomePath\\NameOfYour.dbc;" );


            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "update " + dbFileName + " set zpracovano = "+ zpracovano.ToString()+ " where id_cispod = ?";

                // Now, add the parameters in the same order as the "place-holders" are in above command
                OleDbParameter NewParm = new OleDbParameter("id_cispod", id_cispod);
                NewParm.DbType = DbType.Int32;
                // (or other data type, such as DbType.String, DbType.DateTime, etc)
                cmd.Parameters.Add(NewParm);
                /*
                // Now, on to the next set of parameters...
                NewParm = new OleDbParameter("ParmForAnotherField", NewValueForAnotherField);
                NewParm.DbType = DbType.String;
                MyUpdate.Parameters.Add(NewParm);

                // finally the last one...
                NewParm = new OleDbParameter("ParmForYourKeyField", CurrentKeyValue);
                NewParm.DbType = DbType.Int32;
                MyUpdate.Parameters.Add(NewParm);


                cmd.Parameters.AddWithValue("@var1", id_cispod);
                 */
                cmd.Connection = yourConnectionHandler;
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                string mySQL = @"SELECT cpd, cod, kolik, zpracovano, id_cispod FROM " + dbFileName + " ORDER BY cpd, cod";

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                DA.Fill(ResultSet);
                //DA.Fill(ds);

                /*
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    Console.WriteLine(dr.ItemArray[1].ToString());
                }
                 */
                yourConnectionHandler.Close();
            }

            //return ds;
            return ResultSet;
        }

        /// <summary>
        /// Oznaci nebo odeznaci podnik k razeni 
        /// </summary>
        /// <param name="id_cispod"></param>
        /// <param name="vyrazit"></param>
        /// <returns></returns>
        public DataTable UpdateGRPDataVyrazit(int id_cispod, int vyrazit)
        {
            DataTable ResultSet = new DataTable();
            //DataSet ds = new DataSet();
            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";

            OleDbConnection yourConnectionHandler = new OleDbConnection(
                 //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                 @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            // if including the full dbc (database container) reference, just tack that on
            //      OleDbConnection yourConnectionHandler = new OleDbConnection(
            //          "Provider=VFPOLEDB.1;Data Source=C:\\SomePath\\NameOfYour.dbc;" );


            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "update " + dbFileName + " set Vyrazit = "+ vyrazit.ToString() + " where id_cispod = ?";

                // Now, add the parameters in the same order as the "place-holders" are in above command
                OleDbParameter NewParm = new OleDbParameter("id_cispod", id_cispod);
                NewParm.DbType = DbType.Int32;
                // (or other data type, such as DbType.String, DbType.DateTime, etc)
                cmd.Parameters.Add(NewParm);
                /*
                // Now, on to the next set of parameters...
                NewParm = new OleDbParameter("ParmForAnotherField", NewValueForAnotherField);
                NewParm.DbType = DbType.String;
                MyUpdate.Parameters.Add(NewParm);

                // finally the last one...
                NewParm = new OleDbParameter("ParmForYourKeyField", CurrentKeyValue);
                NewParm.DbType = DbType.Int32;
                MyUpdate.Parameters.Add(NewParm);


                cmd.Parameters.AddWithValue("@var1", id_cispod);
                 */
                cmd.Connection = yourConnectionHandler;
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                string mySQL = @"SELECT vyrazit, cpd, cod, kolik, zpracovano, id_cispod FROM " + dbFileName + " ORDER BY cpd, cod";

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                DA.Fill(ResultSet);
                //DA.Fill(ds);

                /*
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    Console.WriteLine(dr.ItemArray[1].ToString());
                }
                 */
                yourConnectionHandler.Close();
            }

            //return ds;
            return ResultSet;
        }

        public bool UpdateGRPDataVyrazitAll(int vyrazit)
        {
            DataTable ResultSet = new DataTable();
            //DataSet ds = new DataSet();
            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";

            OleDbConnection yourConnectionHandler = new OleDbConnection(
                 //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                 @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            // if including the full dbc (database container) reference, just tack that on
            //      OleDbConnection yourConnectionHandler = new OleDbConnection(
            //          "Provider=VFPOLEDB.1;Data Source=C:\\SomePath\\NameOfYour.dbc;" );


            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "update " + dbFileName + " set Vyrazit = " + vyrazit.ToString() ;
                
                cmd.Connection = yourConnectionHandler;
                cmd.ExecuteNonQuery();
         
                yourConnectionHandler.Close();
            }

            return true;
        }
        public int UpdateDOZData(int id_seznam)
        {
            int kolikZazn = -1;

            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";

            OleDbConnection yourConnectionHandler = new OleDbConnection(
                 @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            string fileName = dbFileName.ToUpper().Replace("GRP_", "");

            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "update " + fileName + " set zpracovano = 1 where id_seznam = ?";

                // Now, add the parameters in the same order as the "place-holders" are in above command
                OleDbParameter NewParm = new OleDbParameter("id_seznam", id_seznam);
                NewParm.DbType = DbType.Int32;
                // (or other data type, such as DbType.String, DbType.DateTime, etc)
                cmd.Parameters.Add(NewParm);
                /*
                // Now, on to the next set of parameters...
                NewParm = new OleDbParameter("ParmForAnotherField", NewValueForAnotherField);
                NewParm.DbType = DbType.String;
                MyUpdate.Parameters.Add(NewParm);

                // finally the last one...
                NewParm = new OleDbParameter("ParmForYourKeyField", CurrentKeyValue);
                NewParm.DbType = DbType.Int32;
                MyUpdate.Parameters.Add(NewParm);

                 */
                try
                {
                    cmd.Connection = yourConnectionHandler;
                    kolikZazn = cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }
                catch
                {
                    kolikZazn = -1;
                }
/*
                string mySQL = @"SELECT cpd, cod, kolik, zpracovano, id_cispod FROM " + dbFileName + " ORDER BY cpd, cod";

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                DA.Fill(ResultSet);
 */ 
                yourConnectionHandler.Close();
            }

            //return ds;
            return kolikZazn;
        }

        public bool KontrolaZpracovaniDOZData(int id_cispod)
        {
            bool vysledek = false;

            DataTable ResultSet = new DataTable();
            //DataSet ds = new DataSet();
            
            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";
            OleDbConnection yourConnectionHandler = new OleDbConnection(
                //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            // if including the full dbc (database container) reference, just tack that on
            //      OleDbConnection yourConnectionHandler = new OleDbConnection(
            //          "Provider=VFPOLEDB.1;Data Source=C:\\SomePath\\NameOfYour.dbc;" );


            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                //string mySQL = @"SELECT * FROM 20141015__46B0JSL4X";  // dbf table name
                string fileName = dbFileName.ToUpper().Replace("GRP_", "");
                string mySQL = @"SELECT COUNT(id_Doz) AS KolikDoz FROM " + fileName + " where id_cispod = ? ";

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                OleDbParameter NewParm = new OleDbParameter("id_cispod", id_cispod);
                NewParm.DbType = DbType.Int32;
                // (or other data type, such as DbType.String, DbType.DateTime, etc)
                MyQuery.Parameters.Add(NewParm);

                OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                DA.Fill(ResultSet);
                /*
                DA.Fill(ds);

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    Console.WriteLine(dr.ItemArray[0].ToString());
                }
                */

                mySQL = @"SELECT COUNT(Zpracovano) AS KolikZprac FROM " + fileName + " where id_cispod = ? AND Zpracovano = 1 ";

                MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                NewParm = new OleDbParameter("id_cispod", id_cispod);
                NewParm.DbType = DbType.Int32;
                // (or other data type, such as DbType.String, DbType.DateTime, etc)
                MyQuery.Parameters.Add(NewParm);

                DA = new OleDbDataAdapter(MyQuery);

                DA.Fill(ResultSet);
                /*
                DA.Fill(ds);

                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    Console.WriteLine(dr.ItemArray[0].ToString());
                }
                 */
                yourConnectionHandler.Close();
            }

            //return ds;

            try
            {
                int KolikDoz = Int32.Parse(ResultSet.Rows[0].ItemArray[0].ToString());
                int KolikZprac = Int32.Parse(ResultSet.Rows[1].ItemArray[1].ToString());
                // vse zpracovano
                if (KolikDoz == KolikZprac)
                    vysledek = true;
            }
            catch
            {
                vysledek = false;
            }            

            return vysledek;
        }

        /// <summary>
        /// Vrati seznam id_Cispod z gridu2, ktere jsou plnnne vyrazene
        /// </summary>
        /// <returns></returns>
        public List<int> DejSeznamPodnikuProVyrazeni()
        {
            DataTable ResultSet = new DataTable();
            List<int> vysledek = new List<int>();
            //DataSet ds = new DataSet();

            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";
            OleDbConnection yourConnectionHandler = new OleDbConnection(
                //@"Provider=VFPOLEDB.1;Data Source=c:\temp\abc\");
                @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            // if including the full dbc (database container) reference, just tack that on
            //      OleDbConnection yourConnectionHandler = new OleDbConnection(
            //          "Provider=VFPOLEDB.1;Data Source=C:\\SomePath\\NameOfYour.dbc;" );


            // Open the connection, and if open successfully, you can try to query it
            yourConnectionHandler.Open();

            if (yourConnectionHandler.State == ConnectionState.Open)
            {
                //string mySQL = @"SELECT * FROM 20141015__46B0JSL4X";  // dbf table name
                string fileGRPName = dbFileName;
                string fileName = dbFileName.ToUpper().Replace("GRP_", "");

                string mySQL = @"SELECT a.id_cispod FROM " + fileName + " a " +
                     " JOIN " + fileGRPName + " b ON a.id_cispod = b.id_cispod " +
                     " where (b.Vyrazit = 1) " +
                     " GROUP BY a.id_cispod";

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                //OleDbParameter NewParm = new OleDbParameter("id_cispod", id_cispod);
                //NewParm.DbType = DbType.Int32;
                //// (or other data type, such as DbType.String, DbType.DateTime, etc)
                //MyQuery.Parameters.Add(NewParm);

                OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                DA.Fill(ResultSet);
                //DA.Fill(ds);

                
                foreach (DataRow dr in ResultSet.Rows)
                {
                    //Console.WriteLine(dr.ItemArray[1].ToString());
                    vysledek.Add(int.Parse(dr.ItemArray[0].ToString()));
                }
                
                yourConnectionHandler.Close();
            }

            //return ds;
            return vysledek;
        }

        /// <summary>
        /// Vypsani udaju z razeneho zaznamu na obrazovku
        /// </summary>
        /// <param name="Tisk_radek_1"></param>
        /// <param name="Tisk_radek_2"></param>
        /// <param name="Tisk_Prijmeni"></param>
        /// <param name="Tisk_cod"></param>
        /// <param name="Tisk_slob"></param>
        /// <param name="Tisk_rok"></param>
        /// <param name="Tisk_mesic"></param>
        /// <returns></returns>
        public bool NastavPopisDoz(string Tisk_radek_1 /*05019017*/, string Tisk_radek_2, string Tisk_Prijmeni, string Tisk_cod, string Tisk_slob, string Tisk_rok, string Tisk_mesic)        
        {
            // 05.04.2016 doplneno tisk COD do eanu, zmena eanu z EAN8 na EAN13
            bool vysledek = false;
            string oddeleni = Tisk_cod; //3
            try
            {

                lblDozimetrRazba.Text = ""; // 06130203
                lblStitekTisk.Text = "";    // 1A_06_130/2_203 Vachata
                lblStitekTiskEan.Text = ""; // 106161302203

                string DozNum = "";
                string DozPopis = "";
                // 05019017;1A Vachata


                // 05019017
                DozNum = Tisk_radek_1.Trim('"', ' ');            // potom bude s COD v retezci, bere se pro tisk stitku
                lblDozNumTab_bezCOD.Text = Tisk_radek_1.Trim('"', ' ');     // bez COD bere se pro razbu dozimetru

                DozNum =  DozNum.Substring(0, 5) + oddeleni +  DozNum.Substring(5, 3);
                //1 Vachata
                //lblDozPopis.Text = DecodeFromUtf8(rowArr[1].Trim('"', ' '));

                if (tisk_z_pole_prijmeni == false)
                {
                    //Pro TiskRadek_2
                    DozPopis = Decodecharset(Tisk_radek_2.Trim('"', ' ')); // pro Tisk_radek_2
                    //1C Vachata
                    DozPopis = DozPopis.Substring(0, 1) +
                                        DejSarziFilmu() +
                                        DozPopis.Substring(1, DozPopis.Length - 1);
                }
                else
                {
                    // pro Tisk_Prijmeni
                    DozPopis = Tisk_radek_2.Trim().Substring(0, 1) + " " + Decodecharset(Tisk_Prijmeni.Trim('"', ' ')); // 1 Vachata
                    //1C Vachata
                    DozPopis = DozPopis.Substring(0, 1) + //1
                                        DejSarziFilmu() +                       //C
                                        DozPopis.Substring(1, DozPopis.Length - 1);
                }                   

                // jak to bude na dozimetru
                string EAN = "";
                string nameZdroj = "";
                string numZdroj = "";
                string namePrvniRadek = ""; 
                string nameDruhyRadek = ""; // Vejsada

                numZdroj =  DozNum.TrimEnd();
                nameZdroj = DozPopis.TrimEnd();

                // 1A_06_130/2_203
                namePrvniRadek = nameZdroj.Substring(0, 2) + '_' + // 1A
                                 numZdroj.Substring(0, 2) + '_' + //  06
                                 numZdroj.Substring(2, 3) + "/" + oddeleni + '_' + // 130/2
                                 numZdroj.Substring(6, 3);   // 203
                // Vachata
                nameDruhyRadek = nameZdroj.Substring(3, nameZdroj.Length - 3); // Vejsada

                // 106151302203
                EAN =            Tisk_slob + // 1
                                 Tisk_mesic + // 06  
                                 Tisk_rok.Substring(2, 2) + // 15
                                 numZdroj.Substring(2, 3) + oddeleni + // 1302
                                 numZdroj.Substring(6, 3);   // 203

                namePrvniRadek = namePrvniRadek.Replace(" ", "");
                nameDruhyRadek = nameDruhyRadek.Replace(" ", "");

                

                //1A Michlova
                //05019001
                //8

                // 2A_MM_PPP_DDD --> 2A_MM_PPPO_DDD
                // Vejsada
                // 0PPPDDD

      
                // 05.04.2016 zmena eanu z EAN8 na EAN13

                lblEANPopis_radek_1.Text = namePrvniRadek;
                lblEANPopis_radek_2.Text = nameDruhyRadek;
                lblDozPopis_radek_1.Text = lblDozNumTab_bezCOD.Text;

                lblDozimetrRazba.Text = lblDozNumTab_bezCOD.Text; // 06130203
                lblStitekTisk.Text = namePrvniRadek + " " + nameDruhyRadek; // 1A_06_130/2_203 Vachata
                lblStitekTiskEan.Text = EAN; // 106161302203

                vysledek = true;

            }
            catch (Exception e)
            {

                lblEANPopis_radek_1.Text = "";
                lblEANPopis_radek_2.Text = "";
                lblDozPopis_radek_1.Text = "";

                lblDozimetrRazba.Text = ""; // 06130203
                lblStitekTisk.Text = "";    // 1A_06_130/2_203 Vachata
                lblStitekTiskEan.Text = ""; // 106161302203

                vysledek = false;
            }
            return vysledek;
        }
  
        /// <summary>
        /// vrati sebrazne stavy zarizeni, chyb atd
        /// </summary>
        public Vlastnosti.popisStavuRaznice DejPopisStavu(bool nelogovat=false)
        {
            Vlastnosti.popisStavuRaznice popisStavu = new Vlastnosti.popisStavuRaznice();
            short nStatus = -1;
            short nInfo = -1;
            short nError = -1;
            popisStavu.stavText = "";
            try
            {
                if (ReadStatus(ref nStatus))
                {
                    popisStavu.nStatusId = nStatus;
                    switch (nStatus)
                    {
                        case 0:
                            popisStavu.nStatusText = "řízení vypnuto";
                            break;
                        case 1:
                            popisStavu.nStatusText = "řízení zapnuto";
                            break;
                        case 2:
                            popisStavu.nStatusText = "automatika zapnuta";
                            break;
                        case 3:
                            popisStavu.nStatusText = "automatika zapnuta a zařízení připraven pro nový příkaz od PC";
                            break;
                        case 4:
                            popisStavu.nStatusText = "chybně zadané parametry, musí se sepnout Reset pro akceptaci chyby";
                            break;
                        case 5:
                            popisStavu.nStatusText = "chyba";
                            break;
                        default:
                            popisStavu.nStatusText = "nedefinováno";
                            nStatus = -1;
                            break;
                    }
                }

                if (ReadInfo(ref nInfo))
                {
                    popisStavu.nInfoId = nInfo;
                    switch (nInfo)
                    {
                        case 0:
                            popisStavu.nInfoText = "Automatický provoz je vypnutý";
                            break;
                        case 1:
                            popisStavu.nInfoText = "Probíhá základní nastavení";
                            break;
                        case 2:
                            popisStavu.nInfoText = "Připraven, čeká na příkaz od PC";
                            break;
                        case 3:
                            popisStavu.nInfoText = "Kontrola příkazu od PC";
                            break;
                        case 4:
                            popisStavu.nInfoText = "Zakládání dílu";
                            break;
                        case 5:
                            popisStavu.nInfoText = "Přesun k zakládání";
                            break;
                        case 6:
                            popisStavu.nInfoText = "Přesun ke kameře";
                            break;
                        case 7:
                            popisStavu.nInfoText = "Kontrola orientace";
                            break;
                        case 8:
                            popisStavu.nInfoText = "Přesun do zmetkovníku";
                            break;
                        case 9:
                            popisStavu.nInfoText = "Přesun k tiskárně";
                            break;
                        case 10:
                            popisStavu.nInfoText = "Tisk";
                            break;
                        case 11:
                            popisStavu.nInfoText = "Přesun k razníku";
                            break;
                        case 12:
                            popisStavu.nInfoText = "Ražení";
                            break;
                        case 13:
                            popisStavu.nInfoText = "Přesun do zásobníku OK dílů";
                            break;
                        case 14:
                            popisStavu.nInfoText = "HOTOVO, přesun do základní polohy";
                            break;
                        case 15:
                            popisStavu.nInfoText = "Řízení vypnuto";
                            break;
                        default:
                            popisStavu.nInfoText = "nedefinováno";
                            popisStavu.nInfoId = -1;
                            break;

                    }
                }

                // jedine v pripade statusu chyba ma smysle cist chybu
                if (popisStavu.nStatusId == 5)
                {
                    if (ReadError(ref nError))
                    {
                        popisStavu.nErrorId = nError;
                        switch (nError)
                        {
                            case 0:
                                popisStavu.nErrorText = "Bez chyby";
                                break;
                            case 1:
                                popisStavu.nErrorText = "Procesorová jednotka zastavena";
                                break;
                            case 8:
                                popisStavu.nErrorText = "Řízení vypnuto";
                                break;
                            case 9:
                                popisStavu.nErrorText = "Ochrany přemostěny";
                                break;
                            case 10:
                                popisStavu.nErrorText = "ESTOP zmáčknut";
                                break;
                            case 11:
                                popisStavu.nErrorText = "Kryt zařízení otevřen";
                                break;
                            case 12:
                                popisStavu.nErrorText = "Nízký tlak";
                                break;
                            case 15:
                                popisStavu.nErrorText = "Nedojel válec – přesun malého založeného dílu z fronty do zařízení (Z20, S10, S11)";
                                break;
                            case 16:
                                popisStavu.nErrorText = "Nedojel válec – přesun velkého založeného dílu z fronty do zařízení (Z21, S12, S13)";
                                break;
                            case 17:
                                popisStavu.nErrorText = "Nedojel válec – zdvih fronty malých OK dílů (Z22, S14, S15)";
                                break;
                            case 18:
                                popisStavu.nErrorText = "Nedojel válec – zdvih fronty velkých OK dílů (Z23, S20, S21)";
                                break;
                            case 19:
                                popisStavu.nErrorText = "Nedojel válec – zdvih vyhazovače NOK dílů (Z24, S22, S23)";
                                break;
                            case 20:
                                popisStavu.nErrorText = "Nedojel válec – vyhazovač NOK dílů (Z25, S24, S25)";
                                break;
                            case 21:
                                popisStavu.nErrorText = "Nedojel válec – otočení tiskové hlavy (Z31, S30, S31)";
                                break;
                            case 23:
                                popisStavu.nErrorText = "Chybně zadané jméno";
                                break;
                            case 24:
                                popisStavu.nErrorText = "Chybně zadané os. číslo";
                                break;
                            case 25:
                                popisStavu.nErrorText = "Chyba v zakládání malého dílu, nezaložen";
                                break;
                            case 26:
                                popisStavu.nErrorText = "Chyba v zakládání velkého dílu, nezaložen";
                                break;
                            case 27:
                                popisStavu.nErrorText = "Vstupní zásobních malých dílů prázdný";
                                break;
                            case 28:
                                popisStavu.nErrorText = "Vstupní zásobních velkých dílů prázdný";
                                break;
                            case 29:
                                popisStavu.nErrorText = "Výstupní zásobních malých dílů plný";
                                break;
                            case 30:
                                popisStavu.nErrorText = "Výstupní zásobních velkých dílů plný";
                                break;
                            case 31:
                                popisStavu.nErrorText = "Chybně zadaný čárový kód";
                                break;
                            case 32:
                                popisStavu.nErrorText = "Chybně zadaný ražený kód";
                                break;
                            case 33:
                                popisStavu.nErrorText = "Chyba v komunikaci s tiskárnou";
                                break;
                            case 34:
                                popisStavu.nErrorText = "Chyba v komunikaci s razníkem";
                                break;
                            case 35:
                                popisStavu.nErrorText = "Zakládání nastavení nedokončeno";
                                break;
                            case 36:
                                popisStavu.nErrorText = "Chyba portálu";
                                break;
                            case 37:
                                popisStavu.nErrorText = "Vložte cartridge CART1 do tiskárny";
                                break;
                            case 38:
                                popisStavu.nErrorText = "Vložte cartridge CART2 do tiskárny";
                                break;
                            case 39:
                                popisStavu.nErrorText = "Vyjměte cartridge z tiskárny";
                                break;
                            default:
                                popisStavu.nErrorText = "nedefinováno";
                                popisStavu.nErrorId = -1;
                                break;

                        }
                    }
                }
                else
                {
                    popisStavu.nErrorId = 0;
                    popisStavu.nErrorText = "Bez chyby";
                }

                popisStavu.stavText = "Status: " + (popisStavu.nStatusText == String.Empty ? "?" : popisStavu.nStatusText) +
                                        ", Info: " + (popisStavu.nInfoText == String.Empty ? "?" : popisStavu.nInfoText) +
                                        ", Error: " + (popisStavu.nErrorText == String.Empty ? "?" : popisStavu.nErrorText);

                if (!nelogovat) 
                    Globalni.Nastroje.LogMessage("DejPopisStavu(), stavText: " + popisStavu.stavText.ToString(), false, "Information", formRaz);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Chyba během volání DejPopisStavu() " + ex.Message.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("DejPopisStavu(), chyba během volání " + ex.Message.ToString(), false, "Error", formRaz);
            }

            return popisStavu;
        }

        /// <summary>
        /// Vyrazeni dozimetru z raziciho planu pres "Vyrazit"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param> 
        private void cmdVyrazit_Click(object sender, EventArgs e)
        {
            bool dorazka = false;
            bool vysledekRaz = false;
            int vyrazenoPocet = 0;
            int rowindexDoz = 0;
            int columnindexDoz = 0;

            toolStripStatusLabel.Text = "";
            lblVyrazenoTab.Text = "0";

            // pro vybrany seznam dozimetru - neorazenych - se provedede orazeni a tisk stitku            
            if (dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("Není co razit ....", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            #region Kontrola
            if (txtSarze.Text.Replace(" ", "") == String.Empty)
            {
                MessageBox.Show("Šarže filmu není zadána", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("cmdVyrazit_Click, Šarže filmu není zadána", false, "Error", formRaz);
                return;
            }
            // kontrola na velikost filmu
            if (txtTyp.Text == String.Empty)
            {
                MessageBox.Show("Typ filmu není zadán", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return;

            }
            if (txtTyp.Text != "1" && txtTyp.Text != "2" && txtTyp.Text != "3")
            {
                MessageBox.Show("Typ filmu není zadán v intervalu 1 - 3", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return;

            }            

            int nTyp = 0;
            if (!int.TryParse(txtTyp.Text, out nTyp))
            {
                MessageBox.Show("Typ filmu není zadán korektně v intervalu 1 - 3", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtTyp.Focus();
                return;

            }
            #endregion

            if (!Init())
            {
                lblStatus.Text = "Chyba komunikace";
                Globalni.Nastroje.LogMessage("cmdVyrazit_Click, Chyba komunikace", false, "Error", formRaz);
                chkReady.Checked = false;
                return;
            }
            chkReady.Checked = true;

            int id_cispod = Int32.Parse(dataGridView2[indexOf(dataGridView2, "Id_Cispod_doz"), 0].Value.ToString());
            Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), dbFileName: " + dbFileName.ToString(), false, "Information", formRaz);
            Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), id_cispod: " + id_cispod.ToString(), false, "Information", formRaz);

            vProcesuRazeni = true;

            // cyklus pres vsechny oznacene filmy k razeni
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                //int col = indexOf(dataGridView2, "Zpracovano_doz");
                //int hodnotaZpracovano = Int32.Parse((row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value).ToString());
                int hodnotaZpracovano = Int32.Parse( (row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value != System.DBNull.Value ? row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value.ToString() : "0"));
                
                if (hodnotaZpracovano == 0) // kdyz je Zpracovano = 0, tak se jeste nerazil dozimetr
                {

                    dataGridView2.Rows[row.Index].Selected = true;
                    dataGridView2.CurrentCell = dataGridView2.Rows[row.Index].Cells[0];

                    if (chkPtatSePredRazbou.Checked == true)
                    {
                        //MessageBox.Show("Razit a tisk dozimetru: " + row.Cells[indexOf(dataGridView2, "Cdz")].Value.ToString());
                        DialogResult result = MessageBox.Show("Ražení a tisk dozimetru č.: "+ (row.Cells[indexOf(dataGridView2, "Cdz")].Value).ToString() +"\r\nPokračovat?", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        switch (result)
                        {
                            case DialogResult.Yes:
                                {
                                    break;
                                }
                            case DialogResult.No:
                                {
                                    vProcesuRazeni = false;
                                    return;
                                    //break;
                                }
                        }
                    }

                    if (row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value == System.DBNull.Value)
                    {
                        dorazka = true;
                        Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), DORazit a DOtisk dozimetru: " + row.Cells[indexOf(dataGridView2, "Cdz")].Value.ToString(), false, "Information", formRaz);
                    }
                    else
                    {
                        dorazka = false;
                        Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), Razit a tisk dozimetru: " + row.Cells[indexOf(dataGridView2, "Cdz")].Value.ToString(), false, "Information", formRaz);
                    }

                    string Tisk_radek_1 = (row.Cells[indexOf(dataGridView2, "Tisk_radek_1")].Value).ToString();
                    string Tisk_radek_2 = (row.Cells[indexOf(dataGridView2, "Tisk_radek_2")].Value).ToString();
                    string Tisk_prijmeni = (row.Cells[indexOf(dataGridView2, "PRIJMENI")].Value).ToString();
                    // 05.04.2016 doplneno tisk COD do eanu, zmena eanu z EAN8 na EAN13
                    string Tisk_cod = (row.Cells[indexOf(dataGridView2, "Oddeleni")].Value).ToString();
                    string Tisk_slob = (row.Cells[indexOf(dataGridView2, "SLOB")]).Value.ToString();
                    string Tisk_rok = (row.Cells[indexOf(dataGridView2, "RP_ROK")]).Value.ToString();
                    string Tisk_mesic = (row.Cells[indexOf(dataGridView2, "RP_MESIC")]).Value.ToString();


                    NastavPopisDoz(Tisk_radek_1, Tisk_radek_2, Tisk_prijmeni, Tisk_cod, Tisk_slob, Tisk_rok, Tisk_mesic);

                    // poslu na raznici a do tisku
                    // parametry v cyklu pres tabulku nepouzivam, hodnoty si zjistim az v telu procedury
                    vysledekRaz = NaRazitDozV2(txt_numDoz: "", txt_nameZdroj: "" , txt_numZdroj: "", txtTyp.Text.ToString());

                    if (!vysledekRaz)
                    {
                        MessageBox.Show("Chyba při ražení dozimetru [" + Tisk_radek_1.TrimEnd() + "] - cyklus ražení byl ukončen.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Globalni.Nastroje.LogMessage("cmdVyrazit_Click, Chyba při ražení dozimetru [" + Tisk_radek_1.TrimEnd() + "] - cyklus ražení byl ukončen.", false, "Error", formRaz);
                        vProcesuRazeni = false;
                        return;
                    }

                    int hodnotaId_Seznam = Int32.Parse((row.Cells[indexOf(dataGridView2, "Id_seznam")].Value).ToString());
                    // kdyz dopadne razeni a tisk, tak zaznam na Zpracovano = 1
                    // pokud budou vsechny dozimetry v pod/odd na Zpracovano = 1, pak i podnik na Zpracovano = 1
                                 
                    if (UpdateDOZData(hodnotaId_Seznam) > 0)
                    {
                        // ukazat zpet v radku zmenu Zpracovano = 1
                        row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value = 1;
                        vyrazenoPocet = vyrazenoPocet + 1;
                        Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), UpdateDOZData(hodnotaId_Seznam): " + hodnotaId_Seznam.ToString(), false, "Information", formRaz);
                    }
                    else
                    {
                        MessageBox.Show("Chyba pri update dozimetru: " + row.Cells[indexOf(dataGridView2, "Cdz")].Value.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), UpdateDOZData(hodnotaId_Seznam): " + hodnotaId_Seznam.ToString(), false, "Error", formRaz);
                    }

                    lblVyrazenoTab.Text = vyrazenoPocet.ToString();

                    // posledni radek, ktery se vyrazil
                    rowindexDoz = dataGridView2.CurrentCell.RowIndex;
                    columnindexDoz = dataGridView2.CurrentCell.ColumnIndex;

                    dataGridView2.Refresh();
                }
            }

            if (dataGridView2.Rows.Count == vyrazenoPocet)
            {
                MessageBox.Show("Vše vyraženo ....", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


            // po orazeni vsech dozimetru 
            // pokud jsou vsechny dozimetry v pod/odd na Zpracovano = 1, pak i podnik na Zpracovano = 1
            int rowindex = dataGridView1.CurrentCell.RowIndex;
            int columnindex = dataGridView1.CurrentCell.ColumnIndex;

            // seznam id_cispod v gridu2
            List<int> seznamId_Cispod = DejSeznamPodnikuProVyrazeni();

            foreach (int id in seznamId_Cispod)
            {
                bool testZpracovaniVsechDoz = KontrolaZpracovaniDOZData(id);
                if (testZpracovaniVsechDoz == true)
                {
                    Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), testZpracovaniVsechDoz pro Id_Cispod = "+ id.ToString()+ ": OK", false, "Information", formRaz);
                    // podnik na  Zpracovano = 1 nastavit
                    DataTable ResultSetUpdate = UpdateGRPData(id, zpracovano: 1);
                }
                else
                    Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), testZpracovaniVsechDoz pro Id_Cispod = " + id.ToString() + ": NE", false, "Information", formRaz);
            }


            dataGridView1.DataSource = "";
            DataTable ResultSet = GetGRPData();
            NastavDataGrid(dataGridView1);
            dataGridView1.DataSource = ResultSet;

            dataGridView1.Rows[rowindex].Selected = true;
            dataGridView1.CurrentCell = dataGridView1.Rows[rowindex].Cells[0];

            dataGridView2.DataSource = "";
            DataTable ResultSetCDZ = GetDOZData();
            NastavDataGrid(dataGridView2);
            dataGridView2.DataSource = ResultSetCDZ;


            try
            {
                dataGridView2.Rows[rowindexDoz].Selected = true;
                dataGridView2.CurrentCell = dataGridView2.Rows[rowindexDoz].Cells[0];
            }
            catch
            {
                if (dataGridView2.Rows.Count != 0)
                {
                    dataGridView2.Rows[0].Selected = true;
                    dataGridView2.CurrentCell = dataGridView2.Rows[0].Cells[0];
                }
            }

            vProcesuRazeni = false;
        }

        private void NastavDataGrid(DataGridView dgv)
        {
            switch (dgv.Name)
            {
                case "dataGridView1":
                    {
                        dataGridView1.Columns[0].DataPropertyName = "Vyrazit"; // checkbox
                        dataGridView1.Columns[1].DataPropertyName = "CPD";
                        dataGridView1.Columns[2].DataPropertyName = "COD";
                        dataGridView1.Columns[3].DataPropertyName = "kolik";
                        dataGridView1.Columns[4].DataPropertyName = "zpracovano"; // checkbox
                        dataGridView1.Columns[5].DataPropertyName = ""; // command button
                        dataGridView1.Columns[6].DataPropertyName = "id_cispod";
                        break;
                    }
                case "dataGridView2":
                    {
                        dataGridView2.Columns[0].DataPropertyName = "CPD";
                        dataGridView2.Columns[1].DataPropertyName = "COD"; //Oddeleni
                        dataGridView2.Columns[2].DataPropertyName = "CDZ";
                        dataGridView2.Columns[3].DataPropertyName = "PRIJMENI";
                        dataGridView2.Columns[4].DataPropertyName = "Tisk_1";
                        dataGridView2.Columns[5].DataPropertyName = "Tisk_2";
                        dataGridView2.Columns[6].DataPropertyName = "zpracovano"; // checkbox
                        dataGridView2.Columns[7].DataPropertyName = "id_cispod"; // ID_Cispod_Doz
                        dataGridView2.Columns[8].DataPropertyName = "id_seznam"; // 
                        dataGridView2.Columns[9].DataPropertyName = "SLOB"; // 
                        dataGridView2.Columns[10].DataPropertyName = "RP_ROK"; // 
                        dataGridView2.Columns[11].DataPropertyName = "RP_MESIC"; // 
                        break;
                    }
            }

        }

        private void cmdOznacitVse_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                //dataGridView2.Rows[row.Index].Selected = true;
                //dataGridView2.CurrentCell = dataGridView2.Rows[row.Index].Cells[0];

                row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value = 1;
            }
        }

        private void cmdOdeznacitVse_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                //dataGridView2.Rows[row.Index].Selected = true;
                //dataGridView2.CurrentCell = dataGridView2.Rows[row.Index].Cells[0];

                row.Cells[indexOf(dataGridView2, "Zpracovano_doz")].Value = 0;
            }
        }
#endregion
        private void Cekej(int seconds)
        {
            Globalni.Nastroje.LogMessage("Cekej, seconds: " + seconds.ToString(), false, "Information", formRaz);
            DateTime Tthen = DateTime.Now;
            do
            {
                Application.DoEvents();
            } while (Tthen.AddSeconds(seconds) > DateTime.Now);  
    
        }

         /// <summary>
        /// Vyrazeni dozimetru na razniciV2 a tisk 
        /// </summary>
        /// <param name="txt_numDoz"> Cislo dozimetru </param>
        /// <param name="txt_nameZdroj"> Stitek horni popisek</param>
        /// <param name="txt_numZdroj"> EAN</param>
        /// <param name="typeDoz"> Typ dozimetru</param>
        /// <returns></returns>
        public bool NaRazitDozV2(string txt_numDoz /*cislo dozimetru*/, string txt_nameZdroj /*popisek stitek*/, string txt_numZdroj /*EAN*/,  string typeDoz /*1,2,3*/)
        {
            // vyrazeni a tisk jednoho dozimetru
            bool vysledekSendText = false;
            bool vysledekFinish = false;
            bool jakSendText = false;

            int kolikrat = 0;
            bool koleckoSendText = false;
            string nameZdroj = ""; // Stitek horni
            string numZdroj = "";  // EAN  
            string numDoz = "";  // cislo dozimetru 
            string namePrint = "";
            string personalNoPrint = "";
            Vlastnosti.popisStavuRaznice popisStavuRaznice = null;
            bool lOk = false;
            int koleckoFinish = 0;

            Globalni.Nastroje.LogMessage("NaRazitDozV2(), start", false, "Information", formRaz);
            if (chkRazitDozimetryTab.Checked == true || chkRazitDozimetry.Checked == true) // na TAB a TAB3
            {
                Globalni.Nastroje.LogMessage("NaRazitDozV2(), start koleckoSendText", false, "Information", formRaz);
                // pokus nekolikrat za sebou
                while (!koleckoSendText)
                {                    
                    kolikrat++;
                    vysledekSendText = true;

                    Globalni.Nastroje.LogMessage("NaRazitDozV2(), kolikrat koleckoSendText: " + kolikrat.ToString() + "x ", false, "Information", formRaz);

                    // a zkola ven, neuspech ?
                    if (kolikrat > koleckoFinishKolikrat)
                    {
                        koleckoSendText = false;
                        break;
                    }

                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu(nelogovat: true);
                    Globalni.Nastroje.LogMessage("NaRazitDozV2(), ReadStatus() [pred popisy]: " + popisStavuRaznice.stavText.ToString(), false, "Information", formRaz);
                    if ((popisStavuRaznice.nStatusId != 3)) //neni chyba, neni řízení vypnuto
                    {
                        Globalni.Nastroje.LogMessage("NaRazitDozV2(), popisStavuRaznice.nStatusId != 3", false, "Information", formRaz);
                        Cekej(1);
                        continue;
                    }



                    // popisky dozimetru
                    if (txt_numDoz == String.Empty && txt_nameZdroj == String.Empty && txt_numZdroj == String.Empty)
                    {
                        // z tisku planu TAB1 
                        nameZdroj = lblStitekTisk.Text.ToString().Trim();    // Stitek horni
                        numZdroj = lblStitekTiskEan.Text.ToString().Trim();  // EAN  
                        numDoz = lblDozimetrRazba.Text.ToString().Trim();    // cislo dozimetru 
                    }
                    else
                    {
                        // postupny TAB2 nebo soubor TAB3, bere se z parametru
                        nameZdroj = txt_nameZdroj.ToString().Trim();    // Stitek horni
                        numZdroj = txt_numZdroj.ToString().Trim();  // EAN  
                        numDoz = txt_numDoz.ToString().Trim();    // cislo dozimetru 
                    }
                    // jmeno a cislo
                    var rows = nameZdroj.Split(' ');
                    if (rows != null)
                    {
                        personalNoPrint = rows[0];
                        namePrint = rows[1];
                    }

                    #region metody SendText
                    // radsi kontrola na max delku znaku
                    /*
                    SendType
                        - nastavení typu dozimetru, na který se bude tisknout(1 = malý, 2 = velký, 3 = velký s otočeným tiskem)
                    SendTextName
                        - nastavení tisknutého jména, maximální délka jména je omezena na 14 znaků
                    SendTextPersonalNo
                        - nastavení tisknutého osobního čísla, délka je pevně stanovena na 15 znaků
                    SendTextBarCode
                        - nastavení tisknutého čárového kódu, délka je pevně stanovena na 12 znaků, 13.znak – kontrolní součet je vypočítán raznicí
                    SendTextRazNo
                        - nastavení raženého kódu, délka je závislá na zvoleném typu dozimetru, 8 znaků pro typ 2 a 3, 6 znaků pro typ 1
                    */
                    if (namePrint.Length > 14)
                    {
                        namePrint = namePrint.Substring(0, 13);
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, pro SendTextName() namePrint zkraceno na 14 znaku", false, "Information", formRaz);
                    }
                    if (personalNoPrint.Length > 15)
                    {
                        personalNoPrint = personalNoPrint.Substring(0, 14);
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, pro SendTextPersonalNo() personalNoPrint zkraceno na 15 znaku", false, "Information", formRaz);
                    }
                    if (numZdroj.Length > 12)
                    {
                        numZdroj = numZdroj.Substring(0, 11);
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, pro SendTextBarCode() numZdroj zkraceno na 12 znaku", false, "Information", formRaz);
                    }

                    if (typeDoz == "1")
                    {
                        if (numDoz.Length > 6)
                        {
                            numDoz = numDoz.Substring(0, 5);
                            Globalni.Nastroje.LogMessage("NaRazitDozV2, pro SendTextRazNo() numZdroj zkraceno na 6 znaku", false, "Information", formRaz);
                        }
                    }
                    else
                    {
                        if (numDoz.Length > 8)
                        {
                            numDoz = numDoz.Substring(0, 7);
                            Globalni.Nastroje.LogMessage("NaRazitDozV2, pro SendTextRazNo() numZdroj zkraceno na 8 znaku", false, "Information", formRaz);
                        }
                    }

                    jakSendText = true;

                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendType(): typeDoz: " + typeDoz.ToString(), false, "Information", formRaz);
                    if (!SendType(typeDoz.ToString()))
                    {
                        popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                        popisStavuRaznice = DejPopisStavu();
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, SendType(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                        jakSendText = false;
                    }
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendType(): namePrint: " + namePrint.ToString(), false, "Information", formRaz);
                    if (!SendTextName(namePrint, namePrint.Length))
                    {
                        popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                        popisStavuRaznice = DejPopisStavu();
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextName(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                        jakSendText = false;
                    }
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextPersonalNo(): personalNoPrint: " + personalNoPrint.ToString(), false, "Information", formRaz);
                    if (!SendTextPersonalNo(personalNoPrint, personalNoPrint.Length))
                    {
                        popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                        popisStavuRaznice = DejPopisStavu();
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextPersonalNo(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                        jakSendText = false;
                    }
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextBarCode(): numZdroj: " + numZdroj.ToString(), false, "Information", formRaz);
                    if (!SendTextBarCode(numZdroj, numZdroj.Length))
                    {
                        popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                        popisStavuRaznice = DejPopisStavu();
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextBarCode(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                        jakSendText = false;
                    }
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextRazNo(): numDoz: " + numDoz.ToString(), false, "Information", formRaz);
                    if (!SendTextRazNo(numDoz, numDoz.Length))
                    {
                        popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                        popisStavuRaznice = DejPopisStavu();
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, SendTextRazNo(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                        jakSendText = false;
                    }
                    #endregion

                    // v pripade neuspechu volani dilcich metod SendText* jdu na zacatek while
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, jakSendText: "+ (jakSendText?"true":"false"), false, "Information", formRaz);
                    if (!jakSendText)
                    {                        
                        vysledekSendText = false;
                        Cekej(1);
                        continue;
                    }

                    Globalni.Nastroje.LogMessage("NaRazitDozV2, Start()", false, "Information", formRaz);
                    if (!Start())
                    {
                        popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                        popisStavuRaznice = DejPopisStavu();
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, Start(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                        vysledekSendText = false;
                        Cekej(1);
                        continue;
                    }

                    // definovane zdrzeni 1sec
                    Cekej(cekejPoStart);

                    // v pripade stavu nStatusId == 4 se vola Reset() a znovu
                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu(nelogovat: true);
                    Globalni.Nastroje.LogMessage("NaRazitDozV2(), ReadStatus() [po Start()]: " + popisStavuRaznice.stavText.ToString(), false, "Information", formRaz);

                    if (popisStavuRaznice.nStatusId == 4)
                    {
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, po Start() nStatusId == 4: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, Reset()", false, "Information", formRaz);

                        if (!Reset())
                        {
                            popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                            popisStavuRaznice = DejPopisStavu();
                            Globalni.Nastroje.LogMessage("NaRazitDozV2, Reset(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);

                            vysledekSendText = false;
                            koleckoSendText = true;
                            break; // a ven z cyklu: while (!vysledek)
                        }

                        // definovane zdrzeni 1sec
                        Cekej(cekejPoReset);
                        continue;
                    }
                    else
                    if (popisStavuRaznice.nStatusId == 3)
                    {
                        break;
                    }

                } // while (!koleckoSendText)
                Globalni.Nastroje.LogMessage("NaRazitDozV2, stop koleckoSendText", false, "Information", formRaz);

                Globalni.Nastroje.LogMessage("NaRazitDozV2, vysledekSendText: " + (vysledekSendText?"true":"false"), false, "Information", formRaz);
                if (vysledekSendText)
                {
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, start koleckoFinish", false, "Information", formRaz);
                    // tady uz mam naslapnuto na uspech, cekam az dojede film na konec ...
                    while (koleckoFinish <= koleckoFinishKolikrat)
                    {
                        Globalni.Nastroje.LogMessage("NaRazitDozV2(), kolikrat koleckoFinish: " + kolikrat.ToString() + "x ", false, "Information", formRaz);
                        // definovane zdrzeni 2sec
                        Cekej(cekejPredReadFinishOK);
                        vysledekFinish = true;

                        lOk = false;
                        Globalni.Nastroje.LogMessage("NaRazitDozV2, ReadFinishOK()", false, "Information", formRaz);
                        if (!ReadFinishOK(ref lOk))
                        {
                            popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                            popisStavuRaznice = DejPopisStavu();
                            Globalni.Nastroje.LogMessage("NaRazitDozV2, ReadFinishOK(): chyba, stav: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                            vysledekFinish = false;
                            koleckoFinish++;
                            continue;
                        }

                        Globalni.Nastroje.LogMessage("NaRazitDozV2, ReadFinishOK() lOk: " + ((lOk)?"true":"false"), false, "Information", formRaz);
                        if (lOk == false)
                        {
                            // jeste neni konec tisku v raznici ...
                            popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                            popisStavuRaznice = DejPopisStavu();

                            // pokud je bez chyby, znovu
                            if (popisStavuRaznice.nStatusId == 5)
                            {
                                Globalni.Nastroje.LogMessage("NaRazitDozV2, ReadFinishOK: lOk = false, popisStavuRaznice.nStatusId == 5: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);

                                // ctu error, ten ale mam uz nacteny
                                Globalni.Nastroje.LogMessage("NaRazitDozV2, ReadFinishOK: lOk = false, popisStavuRaznice.nErroId: " + popisStavuRaznice.nErrorId.ToString() + " -" + popisStavuRaznice.nErrorText.ToString(), false, "Error", formRaz);
                                vysledekFinish = false;
                                break; // z cyklu: while(koleckoFinish <= 3) 
                            }
                            else
                            {
                                // neni chyba, tak znovu
                                Globalni.Nastroje.LogMessage("NaRazitDozV2, ReadFinishOK: lOk = false, popisStavuRaznice: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);

                                // tady jako co?
                                // je to OK, ale co kdyz se opakuje porad? Radsi nastavime neuspech, na zacatku by se melo nastavit na uspech
                                // pokud je ale kolecek hodne, spadne do chyby ...
                                vysledekFinish = false;
                                koleckoFinish++;
                                continue;
                            }
                        }
                        else
                        {
                            // je finis OK, mam narazeno a vytisklo, jdu ven
                            Globalni.Nastroje.LogMessage("NaRazitDozV2, ReadFinishOK: lOk = true, popisStavuRaznice: " + popisStavuRaznice.stavText.ToString(), false, "Information", formRaz);
                            break;
                        }

                    } //while(koleckoFinish <= 3)
                    Globalni.Nastroje.LogMessage("NaRazitDozV2, stop koleckoFinish", false, "Information", formRaz);
                }


            }
            else
            {
                vysledekSendText = true;
                vysledekFinish = true;
            }


            if (vysledekSendText && vysledekFinish)
            {
                // zmacnuti STOP cudle
                if (!vProcesuRazeni)
                {
                    toolStripStatusLabel.Text = "Zachyceni pozadavku STOP";
                    Globalni.Nastroje.LogMessage("NaRazitDozV2(), Zachyceni pozadavku STOP", false, "Information", formRaz);
                    return false;
                }
            }
            else
            {
                toolStripStatusLabel.Text = "Chyba při ražení filmu: " + popisStavuRaznice.stavText.ToString();
                Globalni.Nastroje.LogMessage("NaRazitDozV2(), Chyba při ražení filmu, popisStavuRaznice: " + popisStavuRaznice.stavText.ToString() + ", doz: " + lblDozimetrRazba.Text.ToString(), false, "Error", formRaz);
            }


            Globalni.Nastroje.LogMessage("NaRazitDozV2(), stop: " + ((vysledekSendText && vysledekFinish) ? "true" : "false"), false, "Information", formRaz);
            return (vysledekSendText && vysledekFinish);
        }

        private void cmdOznacitVseVyrazit_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells[indexOf(dataGridView1, "Vyrazit")].Value = 1;
            }

            if (dataGridView1.Rows.Count > 0)
                UpdateGRPDataVyrazitAll(vyrazit: 1);
        }

        private void cmdOdeznacitVseVyrazit_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells[indexOf(dataGridView1, "Vyrazit")].Value = 0;
            }

            if (dataGridView1.Rows.Count > 0)
                UpdateGRPDataVyrazitAll(vyrazit: 0);
        }

        private void cbTypFilmu_SelectedIndexChanged(object sender, EventArgs e)
        {
            short typFilmu = 0;
            Item itm = (Item)cbTypFilmu.SelectedItem;
            typFilmu = (short)itm.Value;
            txtTyp.Text = typFilmu.ToString();
        }
    }
}
