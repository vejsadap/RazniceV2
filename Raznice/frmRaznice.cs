#define NODLL
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
        int DozCount;               // počet vyražených dozimetrů
        int DozMaxCount;            // počet dozimetrů, které se mají orazit
        int DozNum;                 // číslo dozimetru
        string[] DozStr;            // pole čísel dozimetrů
        bool DozFile;               // indikace druhu ražení
        int DozPozice;              // od ktere pozice - radku v souboru se ma razit
        int DozVyrazeno;            // pocitadlo vyrazenych doz v pripade razeni od - do
        string dbFileName = "";     // jmeno dbf souboru s razicim planem
        frmRaznice formRaz;         // nastaveny object formu ve Form_Load pro parametr zasilani LogMessage
        bool tisk_z_pole_prijmeni;  // jake pole (prijmeni nebo Tisk_2) z tabulky se pouzije pro tisk stitku


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
        /// vrati sebrazne stavy zarizeni, chyb atd
        /// </summary>
        public Vlastnosti.popisStavuRaznice DejPopisStavu()
        {
            Vlastnosti.popisStavuRaznice popisStavu = new Vlastnosti.popisStavuRaznice();
            int nStatus = -1;
            int nInfo = -1;
            int nError = -1;
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
                            break;

                    }
                }

                if (ReadError(ref nError))
                {
                    popisStavu.nErrorId = nError;
                    switch (nError)
                    {
                        case 0:
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
                            break;

                    }
                }

                popisStavu.stavText = "Status: " + (popisStavu.nStatusText == String.Empty ? "?" : popisStavu.nStatusText) +
                                        ", Info: " + (popisStavu.nInfoText == String.Empty ? "?" : popisStavu.nInfoText) +
                                        ", Error: " + (popisStavu.nErrorText == String.Empty ? "?" : popisStavu.nErrorText)

            }
            catch(Exception ex)
            {
                MessageBox.Show("Chyba během volání DejPopisStavu() "+ ex.Message.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("DejPopisStavu(), chyba během volání " + ex.Message.ToString(), false, "Error", formRaz);
            }

            return popisStavu;
        }

        public void StartN()
        {
            Vlastnosti.popisStavuRaznice popisStavuRaznice = new Vlastnosti.popisStavuRaznice(); 

            if (txtSarze.Text.Replace(" ", "") == String.Empty)
            {
                MessageBox.Show("Šarže filmu není zadána", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("StartN(), Šarže filmu není zadána", false, "Error", formRaz);
                return;
            }

            bool ok = Init();
            if (!ok) 
            {
                popisStavuRaznice = DejPopisStavu();

                lblStatus.Text = "Chyba komunikace: " + popisStavuRaznice.stavText.ToString();
                Globalni.Nastroje.LogMessage("Init(), Chyba komunikace: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);                
            }

            popisStavuRaznice = DejPopisStavu();
            if ((popisStavuRaznice.nStatusId == 5) || (popisStavuRaznice.nStatusId == 0)) //chyba, řízení vypnuto
            {
                MessageBox.Show("Raznice není připravena: " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("Init(), Raznice není připravena: " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                return;
            }

            string txt = "";
            DozCount = 0;
            lblCount.Text = "0";
            lblCount2.Text = "0";            

            if (DozFile)
            {

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
                            locDozPozice = DozPozice-1;
                            if (locDozPozice < 0)
                                locDozPozice = 0;
                        }
                        else
                            return;
    
                             
                    RozeberDozStr(locDozPozice);
                    DozMaxCount = DozStr.Length;
                    txt = lblDozNum.Text;

                }
                else 
                { 
                    MessageBox.Show("Nelze načíst soubor", Globalni.Parametry.aplikace.ToString(),MessageBoxButtons.OK,MessageBoxIcon.Error);
                    Globalni.Nastroje.LogMessage("StartN(), Nelze načíst soubor: " + txtFile.Text.ToString(), false, "Error", formRaz);
                    return; 
                }
            }
            else
            {
                try
                {
                    DozNum = Convert.ToInt32(txtText.Text);
                    DozMaxCount = Convert.ToInt32(txtCount.Text);
                }
                catch 
                {
                    MessageBox.Show("Číslo je ve špatném formátu", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Globalni.Nastroje.LogMessage("StartN(), Číslo je ve špatném formátu", false, "Error", formRaz);
                    return; 
                }
                txt = InsertSpace(DozNum.ToString());
            }

            // prvni doz, dalsi se resi pres timer2 ...

            if (DozFile)
            {
                // razeni seznamu 
                bool vysledek = false;
                bool jaktisk = false;

                 // tiskne se vse nebo jenom podmnozina ?
                 if ((txtRazitOdDoz.Text.Replace(" ", "").Trim().Length > 0))
                 {
                     // tisknu az z timeru2
                     System.Threading.Thread.Sleep(1000);
                     timer2.Enabled = true;
                 }
                 else
                 {
                    #region razba a pak tisk dozimetru OLD
                    //if (chkRazitDozimetry.Checked == true)
                    // {
                    //     Globalni.Nastroje.LogMessage("StartN(), StartText(txt, txt.Length): " + txt.ToString(), false, "Error", formRaz);
                    //     vysledek = StartText(txt, txt.Length);

                    // }
                    // else
                    //     vysledek = true;

                    //if (vysledek == true)
                    //{
                    //    if (chkTiskSoubor.Checked == true)
                    //    {
                    //        string numZdroj = lblDozNum.Text.ToString().Trim();
                    //        string nameZdroj = lblDozPopis.Text.ToString().Trim();
                    //        string nameZdrojEAN = lblDozPopisEAN.Text.ToString().Trim();

                    //        //jaktisk = Tisk(nameZdroj, numZdroj, false, true);
                    //        jaktisk = Tisk(nameZdroj, nameZdrojEAN, false, true);
                    //    }
                    //    else
                    //        jaktisk = true;

                    //    if (jaktisk == true)
                    //    {
                    //        System.Threading.Thread.Sleep(1000);
                    //        timer2.Enabled = true;
                    //    }
                    //}
                    #endregion

                    // ponovu se nastavi co tisknout a posle se na razbu kde se zaroven i tiskne
                    #region razba s tiskem dozimetru NEW

                    popisStavuRaznice = new Vlastnosti.popisStavuRaznice();
                    popisStavuRaznice = DejPopisStavu();
                    if ((popisStavuRaznice.nStatusId != 3)) //chyba, řízení vypnuto
                    {
                        MessageBox.Show("StartN(): " + popisStavuRaznice.stavText.ToString(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Globalni.Nastroje.LogMessage("StartN(): " + popisStavuRaznice.stavText.ToString(), false, "Error", formRaz);
                        vysledek = false;
                    }

                    if (chkRazitDozimetry.Checked == true)
                    {
                        Globalni.Nastroje.LogMessage("StartN(), StartText(txt, txt.Length): " + txt.ToString(), false, "Error", formRaz);


                        //string numZdroj = lblDozNum.Text.ToString().Trim();
                        string nameZdroj = lblDozPopis.Text.ToString().Trim();
                        string nameZdrojEAN = lblDozPopisEAN.Text.ToString().Trim();

                        // nastavi se vse potrebne
                        if (SetTiskV2(typeDoz:2, txt, nameZdroj, nameZdrojEAN, false, true))
                        {
                        }
                        


                    }
                    else
                        vysledek = true;


                    #endregion



                }
            }
            else
            {
                // tisk jednoho
                ok = StartText(txt, txt.Length);
                if (!ok) 
                { 
                    lblStatus.Text = "Chyba komunikace";
                    Globalni.Nastroje.LogMessage("StartN(), lblStatus.Text: " + lblStatus.Text.ToString(), false, "Error", formRaz);

                }
                else
                {
                    // spustim cyklus pro dozimetry - at nactenych ze souboru nebo ze zalozky 1
                    timer2.Enabled = true;
                }
            }
        }

        public string ErrString(int Err)
        {
            if ((Err > 200) && (Err < 300)) { Err = 200; }
            if ((Err > 400) && (Err < 500)) { Err = 400; }
            if ((Err > 500) && (Err < 600)) { Err = 500; }

            switch (Err)
            {
                case 100:
                    return "Chyba inicializace motoru";
                case 101:
                    return "Chyba komunikace s motorem";
                case 102:
                    return "Došly dozimetry";
                case 103:
                    return "Raznice není připravena";
                case 104:
                    return "Chyba komunikace s razicí jednotkou";
                case 106:
                    return "Zaseknutý dozimetr";
                case 107:
                    return "Central STOP";
                case 108:
                    return "Uživatelský STOP";
                case 109:
                    return "Zaseknutý píst";
                case 200:
                    return "Chyba zapisování textu do razicí jednotky";
                case 300:
                    return "Chyba vybírání masky v razicí jednotce";
                case 400:
                    return "Watchdog – zaseknutí programu";
                case 500:
                    return "Chyba motoru";
                default:
                    return "";
            }
        }

        public string InsertSpace(string txt)
        {
            return txt;
            /*
            txt = txt.Insert(4, " ");
            txt = txt.Insert(1, " ");
            return txt;
             */ 
        }

        #region ImportDLL

#if DLL
   
        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool Init();

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool Ping();

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool Start();

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool Reset(ref bool Status);

        //////////////////////

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool ReadStatus(ref int nStatus);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool ReadInfo(ref int nInfo);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool ReadError(ref int nError);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool ReadFinishOK(ref int lOK);

        //////////////////////

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendType([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendTextName([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendTextPersonalNo([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendTextBarCode([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen);

        [DllImport("RazniceV2.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        [return: MarshalAs(UnmanagedType.I1)]
        static extern bool SendTextRazNo([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen);
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
        public bool ReadStatus(ref int nStatus)
        {
            nStatus = 0;
            return true;
        }

        public bool ReadInfo(ref int nInfo)
        {
            nInfo = 0;
            return true;
        }

        public bool ReadError(ref int nError)
        {
            nError = 0;
            return true;
        }
        public bool ReadFinishOK(ref bool lOK)
        {
            lOK = true;
            return true;
        }
        ////////////////////////////

        public bool SendType(int typ)
        {
            return true;
        }

        public bool SendTextName([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen)
        {
            //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            Globalni.Nastroje.LogMessage("SendTextName:" + txt.Trim(), false, "Information", formRaz);
            //return true;
            Random random = new Random();
            double r = random.NextDouble();
            //int a = (int)r;
            return (r < 0.5 ? false : true);
        }

        public bool SendPersonalNo([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen)
        {
            //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            Globalni.Nastroje.LogMessage("SendPersonalNo:" + txt.Trim(), false, "Information", formRaz);
            //return true;
            Random random = new Random();
            double r = random.NextDouble();
            //int a = (int)r;
            return (r < 0.5 ? false : true);
        }

        public bool SendTextBarCode([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen)
        {
            //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            Globalni.Nastroje.LogMessage("SendTextBarCode:" + txt.Trim(), false, "Information", formRaz);
            //return true;
            Random random = new Random();
            double r = random.NextDouble();
            //int a = (int)r;
            return (r < 0.5 ? false : true);
        }

        public bool SendTextRazNo([MarshalAs(UnmanagedType.LPStr)] string txt, int nLen)
        {
            //MessageBox.Show("StartText:" + text.Trim(), Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
            Globalni.Nastroje.LogMessage("SendTextRazNo:" + txt.Trim(), false, "Information", formRaz);
            //return true;
            Random random = new Random();
            double r = random.NextDouble();
            //int a = (int)r;
            return (r < 0.5 ? false : true);
        }

#endif

//#if DLL
   
//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool Init();

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool InitIP(byte IP1, byte IP2, byte IP3, byte IP4);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool IsReady(ref bool Status);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool IsDone(ref bool done, ref int Err, ref int Mark);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool Start();

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool StartText([MarshalAs(UnmanagedType.LPStr)] string text, int len);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool Run();

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool Stop();

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)] 
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool SendText([MarshalAs(UnmanagedType.LPStr)] string text, int len);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool Mask([MarshalAs(UnmanagedType.LPStr)] string text, int len);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool PrintCode39([MarshalAs(UnmanagedType.LPStr)] string number, int len, string name, int len2);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]            
//        static extern bool PrintEAN8([MarshalAs(UnmanagedType.LPStr)] string number, int len, string name, int len2);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]            
//        static extern bool PrintEAN13([MarshalAs(UnmanagedType.LPStr)] string number, int len, string name, int len2);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool SetIP(byte IP1, byte IP2, byte IP3, byte IP4);

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool PistonUp();

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool PistonDown();

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool Eject();

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern bool ClearInput();

//        [DllImport("Raznice.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
//        [return: MarshalAs(UnmanagedType.I1)]
//        static extern void Disconnect();
//#endif

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

                this.btnSetIP.Enabled = Vlastnosti.allowEdit;
                this.btnMask.Enabled = Vlastnosti.allowEdit;

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


                cbEAN.Items.Add(new Item("EAN13", 2));
                cbEAN.Items.Add(new Item("EAN8", 1));                
                cbEAN.Items.Add(new Item("EAN39", 3));
                cbEAN.SelectedIndex = 0;

                Globalni.Nastroje.LogMessage("Start", false, "Information", formRaz);
                if (!Init())
                {
                    MessageBox.Show("Chyba při inicializování komunikace s PLC", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Globalni.Nastroje.LogMessage("Chyba při inicializování komunikace s PLC", false, "Error", formRaz);
                    //this.Close();
                }
                else
                {

                    txtRazitDoz.Text = "";
                    txtRazitDoz.PromptChar = ' ';
                    txtRazitDoz.Mask = "000";

                    txtRazitOdDoz.Text = "";
                    txtRazitOdDoz.PromptChar = ' ';
                    txtRazitOdDoz.Mask = "00 000 000";

                    timer1.Enabled = true;
                }
            }
            catch
            {
                MessageBox.Show("Nebyla nalezena knihovna Raznice.dll", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("Nebyla nalezena knihovna Raznice.dll", false, "Error", formRaz);
                this.Close();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globalni.Nastroje.LogMessage("Stop", false, "Information", formRaz);
            timer1.Enabled = false;
            timer2.Enabled = false;
            try
            {
                Disconnect();
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

        private void EnablingR(bool ready)
        {
            btnSendText.Enabled = ready;
            btnStart.Enabled = ready;
            btnStarN.Enabled = ready;
            btnStartFromFile.Enabled = ready;            
            chkReady.Checked = ready;
            // zalozka Z tabulky
            cmdVyrazit.Enabled = ready;
            // nastaveni masky - ready + prava
            btnMask.Enabled = Vlastnosti.allowEdit && ready;
        }

        private void EnablingD(bool ready)
        {
            btnUp.Enabled = ready;
            btnDown.Enabled = ready;
            chkDone.Checked = ready;
            btnLoadFile.Enabled = ready;
            // nastaveni IP - ready + prava
            btnSetIP.Enabled = Vlastnosti.allowEdit && ready;
        }

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

        private void btnStart_Click(object sender, EventArgs e)
        {
            //Vyrazit dozimetr
            if (txtSarze.Text == String.Empty)
            {
                MessageBox.Show("Šarže filmu není zadána", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //04373423
            if (txtText.Text.Length != 8)
            {
                MessageBox.Show("Číslo dozimetru musí být 8 znaků MMPPPDDD [např. 04373123].", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            //string txt = InsertSpace(txtText.Text);
            string txt = txtText.Text.Trim();

            bool ok = StartText(txt, txt.Length);
            if (!ok) 
            {
                MessageBox.Show("Raznice není připravena", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            } 



        }

        private void btnSendText_Click(object sender, EventArgs e)
        {
            //Poslat text
            if (txtSarze.Text == String.Empty)
            {
                MessageBox.Show("Šarže filmu není zadána", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //04373423
            if (txtText.Text.Length != 8)
            {
                MessageBox.Show("Číslo dozimetru musí být 8 znaků MMPPPDDD [např. 04373123].", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //string txt = InsertSpace(txtText.Text);
            string txt = txtText.Text.Trim();

            bool ok = SendText(txt, txt.Length);
            if (!ok) { MessageBox.Show("Chyba komunikace", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // Tisk
            string cislo_ean = "";
            string popisek_stitku = "";

#region kontrola

            if (txtSarze.Text == String.Empty)
            {
                MessageBox.Show("Šarže filmu není zadána", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtSarze.Focus();
                return;

            }

            // kontrola vyplneni
            if (txtObdobi.Text == String.Empty)
            {
                MessageBox.Show("Číslo období musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtObdobi.Focus();
                return;
            }
            if (txtMesic.Text == String.Empty)
            {
                MessageBox.Show("Číslo měsíce musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMesic.Focus();
                return;
            }
            if (txtRok.Text == String.Empty)
            {
                MessageBox.Show("Číslo roku musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtRok.Focus();
                return;
            }
            if (txtPodnik.Text == String.Empty)
            {
                MessageBox.Show("Číslo podniku musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPodnik.Focus();
                return;
            }
            if (txtOddeleni.Text == String.Empty)
            {
                MessageBox.Show("Číslo oddělení podniku musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtOddeleni.Focus();
                return;
            }
            if (txtDozimetr.Text == String.Empty)
            {
                MessageBox.Show("Číslo dozimetru musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtDozimetr.Focus();
                return;
            }

//          if (txtJmeno.Text == String.Empty)
//          {
//              MessageBox.Show("Jméno musí být vyplněno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
//              return;
//          }

            int numero =  0;
            // kotrola na delku
            if ((txtObdobi.Text.Length != 1) || !(int.TryParse(txtObdobi.Text, out numero) ))
            {
                MessageBox.Show("Číslo období musí být vyplněno jednou číslicí.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtObdobi.Focus();
                return;
            }
            if ((txtMesic.Text.Length != 2) || !(int.TryParse(txtMesic.Text, out numero) ))
            {
                MessageBox.Show("Číslo měsíce musí být vyplněno dvěma číslicemi.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtMesic.Focus();
                return;
            }
            if ((txtRok.Text.Length != 2) || !(int.TryParse(txtRok.Text, out numero)))
            {
                MessageBox.Show("Číslo roku musí být vyplněno dvěma číslicemi.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtRok.Focus();
                return;
            }
            if ((txtPodnik.Text.Length != 3) || !(int.TryParse(txtPodnik.Text, out numero)))
            {
                MessageBox.Show("Číslo podniku musí být vyplněno třemi číslicemi.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtPodnik.Focus();
                return;
            }
            if ((txtOddeleni.Text.Length != 1) || !(int.TryParse(txtOddeleni.Text, out numero)))
            {
                MessageBox.Show("Číslo oddělení podniku musí být vyplněno jednou číslicí.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtOddeleni.Focus();
                return;
            }
            if ((txtDozimetr.Text.Length != 3) || !(int.TryParse(txtDozimetr.Text, out numero)))
            {
                MessageBox.Show("Číslo dozimetru musí být vyplněno třemi číslicemi.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtDozimetr.Focus();
                return;
            }

#endregion
            // 1 06 16 130 2 203

            // 1A_06_130/2_203
            popisek_stitku = txtObdobi.Text.Trim() + txtSarze.Text.Trim() +  '_' + // 1A
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

            // tisk popisku z tab. Postupna 
            Tisk(popisek_stitku, cislo_ean, true, false);

        }

        private void btnStarN_Click(object sender, EventArgs e)
        {
            // Vyrazit N dozimetrů
            if (txtText.Text.Length != 8)
            {
                MessageBox.Show("Číslo dozimetru musí být 8 znaků MMPPPDDD [např. 04373123].", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DozFile = false;
            StartN();
        }

        private void btnStartFromFile_Click(object sender, EventArgs e)
        {
            DozFile = true;
            StartN();
        }
        
        private void btnStop_Click(object sender, EventArgs e)
        {
#if DLL
            if (btnStop.Text == "STOP") { Stop(); }
            else { Run(); }
#else
            //simulace fci z Raznice.dll
            timer2.Enabled = false;
            EnablingR(true);

#endif
        }

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

        private void btnMask_Click(object sender, EventArgs e)
        {
            bool ok = Mask(txtMask.Text, txtMask.Text.Length);
            if (!ok) { MessageBox.Show("Chyba komunikace", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void btnSetIP_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            string[] ip = txtIP.Text.Split('.');
            byte[] ipb;
            ipb = new byte[4];

            try
            {
                ipb[0] = Convert.ToByte(ip[0]);
                ipb[1] = Convert.ToByte(ip[1]);
                ipb[2] = Convert.ToByte(ip[2]);
                ipb[3] = Convert.ToByte(ip[3]);
            }
            catch
            {
                MessageBox.Show("Špatně zadaná IP adresa", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            bool ok = SetIP(ipb[0], ipb[1], ipb[2], ipb[3]);
            if (!ok) { MessageBox.Show("Chyba komunikace", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error); }


            timer1.Enabled = true;
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            bool ok = PistonUp();
            if (!ok) { MessageBox.Show("Raznice není připravena", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            bool ok = PistonDown();
            if (!ok) { MessageBox.Show("Raznice není připravena", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void btnEject_Click(object sender, EventArgs e)
        {
            Eject();
        }

        private void btnReconnect_ButtonClick(object sender, EventArgs e)
        {
            btnStop.Enabled = true;
            timer1.Enabled = true;
            btnReconnect.Visible = false;
        }

        private void btnClearInput_Click(object sender, EventArgs e)
        {
            ClearInput();
        }

#endregion

#region Timery

        private void timer1_Tick(object sender, EventArgs e)
        {
            bool ready = false, done = false;
            int Err = 0, Mark = 0;
            bool ok = IsDone(ref done, ref Err, ref Mark);
            if (!ok)
            {
                STPbtn(true);
                lblStatus.Text = "Chyba komunikace";
                timer1.Enabled = false;
                btnStop.Enabled = false;
                btnReconnect.Visible = true;
                EnablingD(false); 
                EnablingR(false); 
                return;
            }
            else
            {
                if (timer2.Enabled) 
                { 
                    EnablingD(false); 
                } 
                else 
                    EnablingD(done);

                if (Err > 0)
                {
                    STPbtn(false);
                    lblStatus.Text = "Error: " + ErrString(Err);
                    
                    switch (Mark)
                    {
                        case 0:
                            lblMark.Text = "- Poslední dozimetr neoražen";                            
                            break;
                        case 1:
                            lblMark.Text = "- Poslední dozimetr oražen";
                            break;
                        case 2:
                            lblMark.Text = "- Nelze zjistit, je li poslední dozimetr oražen správně";
                            break;
                    }
                    Globalni.Nastroje.LogMessage("timer1_Tick, lblStatus.Text " + lblStatus.Text.ToString() + ", lblMark.Text: " + lblMark.Text.ToString(), false, "Error", formRaz);
                }
                else
                {
                    STPbtn(true);
                    lblStatus.Text = "";
                    lblMark.Text = "";
                }
            }

            ok = IsReady(ref ready);
            if (timer2.Enabled) { EnablingR(false); } else EnablingR(ready);
            if (!ok) 
            { 
                lblStatus.Text = "Chyba komunikace";
                Globalni.Nastroje.LogMessage("timer1_Tick, lblStatus.Text " + lblStatus.Text.ToString(), false, "Error", formRaz);
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            bool done = false;
            bool ready = false;
            int err = 0, mark = 0;

            bool ok = IsDone(ref done, ref err, ref mark);
            if (!ok)
            {
                timer2.Enabled = false;
                MessageBox.Show("Ztráta spojení s raznicí", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("timer2_Tick, Ztráta spojení s raznicí", false, "Error", formRaz);
                return;
            }
            if (done)
            {
                if (err == 0)
                {
                    DozCount += 1;
#region DozFile
                    if (DozFile)
                    {
                        // tiskne se vse nebo jenom podmnozina ?
                        if ((txtRazitOdDoz.Text.Replace(" ", "").Trim().Length > 0)
                            &&
                                (!((DozCount >= DozPozice)
                                &&
                                //(DozCount <= (DozPozice + int.Parse(txtRazitDoz.Text.Replace(" ", "").Trim()) - 1))))
                                //(DozCount <= (DozPozice + int.Parse(txtRazitDoz.Text.Replace(" ", "").Trim()) + 1))))
                                (DozCount <= (DozPozice + int.Parse(txtRazitDoz.Text.Replace(" ", "").Trim()) ))))
                            )
                        {
                            if ((err > 0) || (DozCount >= DozMaxCount))
                            {
                                timer2.Enabled = false;
                                return;
                            }
                            // vynechavam razeni, nejsem v intervalu
                            lblStatus.Text = "Skip dozimetru";
                            Globalni.Nastroje.LogMessage("timer2_Tick, Skip dozimetru", false, "Information", formRaz);
                            return;
                        }
                        else
                        {
                            DozVyrazeno += 1;
                        }

                        lblCount2.Text = DozCount.ToString();


                        if (DozCount < DozStr.Length)
                        {
                            RozeberDozStr(DozCount);
                        }
                        else
                        {
                            RozeberDozStr(0);
                        }
/*

                        if (DozCount < DozStr.Length) 
                            lblDozNum.Text = DozStr[DozCount];
                        else 
                            lblDozNum.Text = DozStr[0];
*/
                    }
#endregion
                    else
#region Jednotlive
                    {
                        DozNum += 1;
                        lblCount.Text = DozCount.ToString();
                        txtText.Text = DozNum.ToString();
                    }
#endregion
                }

                if ((err > 0) || (DozCount >= DozMaxCount))
                {
                    timer2.Enabled = false;
                    return;
                }

                string txtDoz = "";
                string txtName = "";
                string txtEAN = "";

                if (DozFile)
                {
                    RozeberDozStr(DozCount);
                    txtDoz = lblDozNum.Text;
                    txtName = lblDozPopis.Text;
                    txtEAN = lblDozPopisEAN.Text;
                }
                else
                {
                    txtDoz = InsertSpace(DozNum.ToString());
                }


                //pokud je zadano omezeni intervalu dozimetru k tisku
                

                int i = 0;
                bool vysledek = false;
                // priznak, ze se ma vubec provadet razeni
                if (chkRazitDozimetry.Checked == true)
                {
                    Globalni.Nastroje.LogMessage("timer2_Tick, StartText(txtDoz, txtDoz.Length)" + txtDoz.ToString(), false, "Information", formRaz);
                    vysledek = StartText(txtDoz, txtDoz.Length);
                }
                else
                    vysledek = true;

                while (!vysledek)
                {
                    vysledek = StartText(txtDoz, txtDoz.Length);
                    i++;
                    if (i > 3)
                    {
                        timer2.Enabled = false;
                        MessageBox.Show("Ztráta spojení s raznicí", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Globalni.Nastroje.LogMessage("timer2_Tick, Ztráta spojení s raznicí", false, "Error", formRaz);
                        return;
                    }
                    System.Threading.Thread.Sleep(100);
                }
                if (vysledek == true)
                {
                    if (chkTiskSoubor.Checked == true)
                    {
                        // Unicode Character 'NO-BREAK SPACE' (U+00A0)
                        //string name = "doc.\u00A0Pepa Novák, CSc.";
                        //string num = "4 948 17";

                        //string name = "";
                        //string num = "";
                        string numZdroj = "";
                        string nameZdroj = "";
                        //string namePrvniRadek = "";
                        //string nameDruhyRadek = "";

                        // 2A_MM_PPP_DDD
                        // Vejsada
                        // 0PPPDDD

                        numZdroj = txtDoz.ToString().Trim();
                        nameZdroj = txtName.ToString().Trim(); 

                        // tisk ze souboru
                        //Tisk(nameZdroj, numZdroj, false, true);                        
                        Tisk(nameZdroj, txtEAN, false, true);


                    }
                }
            }
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

        /// <summary>
        /// Nastaveni textu pro razbu a potisk 
        /// </summary>
        /// <param name="txt_DozNum"></param>
        /// <param name="popisek_stitek"></param>
        /// <param name="cislo_ean"></param>
        /// <param name="hlasitChybu"></param>
        /// <param name="VolnyTisk"></param>
        /// <returns></returns>
        private bool SetTiskV2(int typeDoz /*1,2,3*/, string txt_DozNum /*11001003*/, string popisek_stitek /*1A_06_130/2_203 Michlova*/, string cislo_ean /*106151302203*/, bool hlasitChybu, bool VolnyTisk)
        {
            bool jakTisk = false;
            string nameZdroj = "";
            string numZdroj = "";
            string namePrint = "";
            string personalNoPrint = "";

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
                    personalNoPrint = rows[1];
                    namePrint = rows[2];
                }

                if (!VolnyTisk)
                {
                    if ((numZdroj.Length != 12) && (hlasitChybu))
                        MessageBox.Show("Číslo pro konstrukci EAN kódu: '" + numZdroj.ToString() + "' musí být dlouhé 12 znaků.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (numZdroj.Length != 12)
                    {
                        Globalni.Nastroje.LogMessage("Tisk PrintEAN13, Číslo pro konstrukci EAN kódu musí být dlouhé 12 znaků, numZdroj:" + numZdroj.ToString(), false, "Error", formRaz);
                        return false;
                    }

                    //1A_06_130/2_203
                    if ((nameZdroj.Length < 15) && (hlasitChybu))
                        MessageBox.Show("Text štítku dozimetru '" + numZdroj.ToString() + "' musí být minimálně 15 znaků dlouhý.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // pokud je mensi, doplnim na 15 pozice - je mozne tisknout jen cislo dozimetru beze jmena
                    if (nameZdroj.Length < 15)
                    {
                        Globalni.Nastroje.LogMessage("Tisk PrintEAN13, Text štítku dozimetru musí být minimálně 15 naků dlouhý., numZdroj:" + nameZdroj.ToString(), false, "Warning", formRaz);
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


                Globalni.Nastroje.LogMessage("Tisk PrintEAN13 num:" + numZdroj.ToString() + ", name: " + nameZdroj.ToString(), false, "Information", formRaz);
                //jakTisk = PrintEAN13(numZdroj, numZdroj.Length, nameZdroj, nameZdroj.Length);

                if (SendType(typeDoz))
                    if (SendTextName(namePrint, namePrint.Length))
                        if (SendPersonalNo(personalNoPrint, personalNoPrint.Length))
                            if (SendTextBarCode(numZdroj, numZdroj.Length))
                                if (SendTextRazNo(txt_DozNum, txt_DozNum.Length))
                                    jakTisk = true;

                if (jakTisk == true)
                {
                    lblStatus.Text = "Tisk EAN13 ok";
                    //toolStripStatusLabel.Text = "Tisk EAN13 ok";
                }
                else
                {
                    lblStatus.Text = "Chyba tisku EAN13";
                    toolStripStatusLabel.Text = "Chyba tisku EAN13";
                    if (hlasitChybu)
                        MessageBox.Show("Chyba tisku EAN13", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                Globalni.Nastroje.LogMessage("Tisk PrintEAN13 " + lblStatus.Text, false, "Information", formRaz);
            }
            catch (Exception ex)
            {
                Globalni.Nastroje.LogMessage("Chyba SetTiskV2(): " + ex.Message.ToString(), false, "Error", formRaz);
            }
            return jakTisk;
        }

        private bool Tisk(string popisek_stitek /*1A_06_130/2_203 Michlova*/, string cislo_ean /*106151302203*/, bool hlasitChybu, bool VolnyTisk)
        {
            bool jakTisk = false;
            string nameZdroj = "";
            string numZdroj = "";

            //1A Michlova
            //050190002

            // od 05.04.2016 obsahuje i cislo oddeleni: 0PPPDDD--> 0PPPODDD
            // Vejsada
            // 0PPPDDD


            nameZdroj = popisek_stitek.Trim(); //1A_06_130/2_203 Michlova
            numZdroj = cislo_ean.Trim();       //106151302203

            if (!VolnyTisk)
            {
                if ((numZdroj.Length != 12) && (hlasitChybu))
                    MessageBox.Show("Číslo pro konstrukci EAN kódu: '" + numZdroj.ToString() + "' musí být dlouhé 12 znaků.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (numZdroj.Length != 12)
                {
                    Globalni.Nastroje.LogMessage("Tisk PrintEAN13, Číslo pro konstrukci EAN kódu musí být dlouhé 12 znaků, numZdroj:" + numZdroj.ToString(), false, "Error", formRaz);
                    return false;
                }

                //1A_06_130/2_203
                if ((nameZdroj.Length < 15) && (hlasitChybu))
                    MessageBox.Show("Text štítku dozimetru '" + numZdroj.ToString() + "' musí být minimálně 15 znaků dlouhý.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                // pokud je mensi, doplnim na 15 pozice - je mozne tisknout jen cislo dozimetru beze jmena
                if (nameZdroj.Length < 15)
                {
                    Globalni.Nastroje.LogMessage("Tisk PrintEAN13, Text štítku dozimetru musí být minimálně 15 naků dlouhý., numZdroj:" + nameZdroj.ToString(), false, "Warning", formRaz);
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


            Globalni.Nastroje.LogMessage("Tisk PrintEAN13 num:" + numZdroj.ToString() + ", name: " + nameZdroj.ToString(), false, "Information", formRaz);
            jakTisk = PrintEAN13(numZdroj, numZdroj.Length, nameZdroj, nameZdroj.Length);

            if (jakTisk == true)
            {
                lblStatus.Text = "Tisk EAN13 ok";
                //toolStripStatusLabel.Text = "Tisk EAN13 ok";
            }
            else
            {
                lblStatus.Text = "Chyba tisku EAN13";
                toolStripStatusLabel.Text = "Chyba tisku EAN13";
                if (hlasitChybu)
                    MessageBox.Show("Chyba tisku EAN13", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Globalni.Nastroje.LogMessage("Tisk PrintEAN13 " + lblStatus.Text, false, "Information", formRaz);
            return jakTisk;
        }

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
                if (e.ColumnIndex == indexOf(dataGridView1,"Nacist")) // button Nacist
                    //dataGridView1.Columns[e.ColumnIndex].Name = "Nacist";
                    

                {
                    int zpracovano = Int32.Parse(dataGridView1[indexOf(dataGridView1, "Zpracovano"), e.RowIndex].Value.ToString());
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
                                    break;
                                }
                            case DialogResult.No:
                                {
                                    return;
                                    //break;
                                }
                        }
                    }

                    
                    
                    //int id_cispod = dataGridView1.Columns[5];
                    //DataGridViewRow row = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];

                    int rowindex = dataGridView1.CurrentCell.RowIndex;
                    int columnindex = dataGridView1.CurrentCell.ColumnIndex;

                    string a = dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();

                    int id_cispod = Int32.Parse(dataGridView1[indexOf(dataGridView1, "Id_Cispod"), e.RowIndex].Value.ToString());                    

                    dataGridView1.DataSource = "";
                                       
                    //DataTable ResultSet = UpdateGRPData(id_cispod);                    
                    DataTable ResultSet = GetGRPData();
                    NastavDataGrid(dataGridView1);
                    dataGridView1.DataSource = ResultSet;

                    dataGridView1.Rows[rowindex].Selected = true;
                    dataGridView1.CurrentCell = dataGridView1.Rows[rowindex].Cells[0];

                    dataGridView2.DataSource = "";
                    DataTable ResultSetCDZ = GetDOZData(id_cispod);
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

                string mySQL = @"SELECT cpd, cod, kolik, zpracovano, id_cispod FROM " + dbFileName + " ORDER BY cpd, cod";

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                OleDbDataAdapter DA = new OleDbDataAdapter(MyQuery);

                DA.Fill(ResultSet);

                yourConnectionHandler.Close();
            }

            //return ds;
            return ResultSet;
        }

        public DataTable GetDOZData(int id_cispod)
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
                string fileName = dbFileName.Replace("GRP_", "");
                string mySQL = @"SELECT cpd, cod, Cdz, Prijmeni, Tisk_1, Tisk_2, zpracovano, id_seznam, id_cispod, SLOB, RP_ROK, RP_MESIC FROM " + fileName + " where id_cispod = ? ORDER BY cpd, cod, cdz";

                OleDbCommand MyQuery = new OleDbCommand(mySQL, yourConnectionHandler);
                OleDbParameter NewParm = new OleDbParameter("id_cispod", id_cispod);
                NewParm.DbType = DbType.Int32;
                // (or other data type, such as DbType.String, DbType.DateTime, etc)
                MyQuery.Parameters.Add(NewParm);

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

        public DataTable UpdateGRPData(int id_cispod)
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
                cmd.CommandText = "update " + dbFileName + " set zpracovano = 1 where id_cispod = ?";

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

        public int UpdateDOZData(int id_seznam)
        {
            int kolikZazn = -1;

            string filepath = Path.GetDirectoryName(dbFileName);
            if (!filepath.EndsWith("\\"))
                filepath += "\\";

            OleDbConnection yourConnectionHandler = new OleDbConnection(
                 @"Provider=VFPOLEDB.1;Data Source=" + filepath);

            string fileName = dbFileName.Replace("GRP_", "");

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
                string fileName = dbFileName.Replace("GRP_", "");
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

        public bool NaRazitDoz(string Tisk_radek_1, string Tisk_radek_2)
        {
            // vyrazeni a tisk jednoho dozimetru
            bool vysledek = false;
            bool jaktisk = false;
            int i = 0;

            bool ready = false; 
            bool done = false;
            int Err = 0;
            int Mark = 0;
            int kolikrat = 0;
            bool konecRazeni = false;

            if (chkRazitDozimetryTab.Checked == true)
            {
                // pokus nekolikrat za sebou
                while (!konecRazeni)
                {
                    kolikrat++;
                    Globalni.Nastroje.LogMessage("NaRazitDoz(), kolikrat: " + kolikrat.ToString() + "x ", false, "Information", formRaz);
                    

                    // a zkola ven ?
                    if (kolikrat > 6)
                        konecRazeni = true;

                    // otestuju si, zda se da vubec razit film
                    // 1. je ukoncena razba?
                    bool ok = IsDone(ref done, ref Err, ref Mark);
                    if (!ok)
                    {
                        Globalni.Nastroje.LogMessage("NaRazitDoz(), IsDone: Ztráta spojení s raznicí", false, "Error", formRaz);
                        Cekej(5);
                        continue;
                    }

                    if (done == true)
                    {
                        // 2. je vse ready, pripraveno k dalsi razbe?
                        ok = IsReady(ref ready);                        
                        if (!ok)
                        {
                            Globalni.Nastroje.LogMessage("NaRazitDoz(), IsReady: Chyba komunikace", false, "Error", formRaz);
                            Cekej(5);
                            continue;
                        }

                        if (ready == true)
                        {
                            Globalni.Nastroje.LogMessage("NaRazitDoz(), StartText(lblDozimetrRazba.Text, lblDozimetrRazba.Text.Length): " + lblDozimetrRazba.Text.ToString(), false, "Information", formRaz);
                            vysledek = StartText(lblDozimetrRazba.Text, lblDozimetrRazba.Text.Length);
                        }

                        if (vysledek == true)
                            konecRazeni = true;

                        while (!vysledek)
                        {
                            vysledek = StartText(lblDozimetrRazba.Text, lblDozimetrRazba.Text.Length);
                            i++;
                            if (i > 3)
                            {
                                Globalni.Nastroje.LogMessage("NaRazitDoz(), while (!vysledek), Ztráta spojení s raznicí: " + i.ToString() + "x ", false, "Error", formRaz);
                                vysledek = false;
                                konecRazeni = true;
                                break; // a ven z cyklu: while (!vysledek)
                            }
                            Cekej(5);        

                        }
                        
                    }

                    Cekej(5);
                } // while (!konecRazeni)
            }
            else
                vysledek = true;


            if (vysledek == true)
            {
                if (chkTiskSouborTab.Checked == true)
                {
                    string nameZdroj = lblStitekTisk.Text.ToString().Trim();    // Stitek horni
                    string numZdroj = lblStitekTiskEan.Text.ToString().Trim();  // EAN  
                    

                    Globalni.Nastroje.LogMessage("NaRazitDoz(), Tisk(nameZdroj, numZdroj, false): " + nameZdroj.ToString() + ", " + numZdroj.ToString(), false, "Information", formRaz);
                    jaktisk = Tisk(nameZdroj, numZdroj, false, false);

                    if (jaktisk != true)
                    {
                        Globalni.Nastroje.LogMessage("NaRazitDoz(), Chyba při tisku, Tisk(nameZdroj, numZdroj, false): " + nameZdroj.ToString() + ", " + numZdroj.ToString(), false, "Error", formRaz);
                    }

                }
                else
                    jaktisk = true;
            }
            else
            {
                toolStripStatusLabel.Text = "Chyba při ražení filmu";
                Globalni.Nastroje.LogMessage("NaRazitDoz(), Chyba při ražení filmu, vysledek StartText() == false: " + lblDozimetrRazba.Text.ToString(), false, "Error", formRaz);
            }



            return (vysledek && jaktisk);
        }

        private void cmdVyrazit_Click(object sender, EventArgs e)
        {
            bool dorazka = false;
            bool vysledekRaz = false;
            int vyrazenoPocet = 0;
            int rowindexDoz = 0;
            int columnindexDoz = 0;

            lblVyrazenoTab.Text = "0";

            // pro vybrany seznam dozimetru - neorazenych - se provedede orazeni a tisk stitku            
            if (dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("Není co razit ....", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            if (txtSarze.Text.Replace(" ", "") == String.Empty)
            {
                MessageBox.Show("Šarže filmu není zadána", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("cmdVyrazit_Click, Šarže filmu není zadána", false, "Error", formRaz);
                return;
            }

            bool ready = false;
            bool ok = IsReady(ref ready);
            if (!ok)
            {
                lblStatus.Text = "Chyba komunikace";
                Globalni.Nastroje.LogMessage("cmdVyrazit_Click, Chyba komunikace", false, "Error", formRaz);
            }
            if (!ready)
            {
                MessageBox.Show("Raznice není připravena", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                Globalni.Nastroje.LogMessage("cmdVyrazit_Click, Raznice není připravena", false, "Error", formRaz);
                return;
            }



            int id_cispod = Int32.Parse(dataGridView2[indexOf(dataGridView2, "Id_Cispod_doz"), 0].Value.ToString());
            Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), dbFileName: " + dbFileName.ToString(), false, "Information", formRaz);
            Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), id_cispod: " + id_cispod.ToString(), false, "Information", formRaz);

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
                    vysledekRaz = NaRazitDoz(Tisk_radek_1, Tisk_radek_2);

                    if (!vysledekRaz)
                    {
                        MessageBox.Show("Chyba při ražení dozimetru [" + Tisk_radek_1.TrimEnd() + "] - cyklus ražení byl ukončen.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Globalni.Nastroje.LogMessage("cmdVyrazit_Click, Chyba při ražení dozimetru [" + Tisk_radek_1.TrimEnd() + "] - cyklus ražení byl ukončen.", false, "Error", formRaz);
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
            bool testZpracovaniVsechDoz = KontrolaZpracovaniDOZData(id_cispod);
            if (testZpracovaniVsechDoz == true) 
            {
                Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), testZpracovaniVsechDoz: OK", false, "Information", formRaz);
                // podnik na  Zpracovano = 1 nastavit
                int rowindex = dataGridView1.CurrentCell.RowIndex;
                int columnindex = dataGridView1.CurrentCell.ColumnIndex;

                dataGridView1.DataSource = "";
                DataTable ResultSet = UpdateGRPData(id_cispod);                    
                NastavDataGrid(dataGridView1);
                dataGridView1.DataSource = ResultSet;

                dataGridView1.Rows[rowindex].Selected = true;
                dataGridView1.CurrentCell = dataGridView1.Rows[rowindex].Cells[0];

                dataGridView2.DataSource = "";
                DataTable ResultSetCDZ = GetDOZData(id_cispod);
                NastavDataGrid(dataGridView2);
                dataGridView2.DataSource = ResultSetCDZ;

                //dataGridView2.Rows[0].Selected = true;
               //dataGridView2.CurrentCell = dataGridView1.Rows[0].Cells[0];


            }
            else
                Globalni.Nastroje.LogMessage("cmdVyrazit_Click(), testZpracovaniVsechDoz: NE", false, "Information", formRaz);


            dataGridView2.Rows[rowindexDoz].Selected = true;
            dataGridView2.CurrentCell = dataGridView2.Rows[rowindexDoz].Cells[0];
        }

        private void NastavDataGrid(DataGridView dgv)
        {
            switch (dgv.Name)
            {
                case "dataGridView1":
                    {
                        dataGridView1.Columns[0].DataPropertyName = "CPD";
                        dataGridView1.Columns[1].DataPropertyName = "COD";
                        dataGridView1.Columns[2].DataPropertyName = "kolik";
                        dataGridView1.Columns[3].DataPropertyName = "zpracovano"; // checkbox
                        dataGridView1.Columns[4].DataPropertyName = ""; // command button
                        dataGridView1.Columns[5].DataPropertyName = "id_cispod";
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

        private void btnMask_Click_1(object sender, EventArgs e)
        {
            bool ok = Mask(txtMask.Text, txtMask.Text.Length);
            if (!ok) 
            {
                Globalni.Nastroje.LogMessage("btnMask_Click_1, Chyba komunikace, txtMask: " + txtMask.Text.ToString(), false, "Error", formRaz);
                MessageBox.Show("Chyba komunikace", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
        }

        private void btnSetIP_Click_1(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            string[] ip = txtIP.Text.Split('.');
            byte[] ipb;
            ipb = new byte[4];

            try
            {
                ipb[0] = Convert.ToByte(ip[0]);
                ipb[1] = Convert.ToByte(ip[1]);
                ipb[2] = Convert.ToByte(ip[2]);
                ipb[3] = Convert.ToByte(ip[3]);
            }
            catch
            {
                Globalni.Nastroje.LogMessage("btnSetIP_Click_1, Špatně zadaná IP adresa, txtIP: " + txtIP.Text.ToString(), false, "Error", formRaz);
                MessageBox.Show("Špatně zadaná IP adresa", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            bool ok = SetIP(ipb[0], ipb[1], ipb[2], ipb[3]);
            if (!ok) 
            {
                Globalni.Nastroje.LogMessage("btnSetIP_Click_1, Chyba komunikace, txtIP: " + txtIP.Text.ToString(), false, "Error", formRaz);
                MessageBox.Show("Chyba komunikace", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }


            timer1.Enabled = true;
        }

        private void Cekej(int seconds)
        {
            Globalni.Nastroje.LogMessage("Cekej, seconds: " + seconds.ToString(), false, "Information", formRaz);
            DateTime Tthen = DateTime.Now;
            do
            {
                Application.DoEvents();
            } while (Tthen.AddSeconds(seconds) > DateTime.Now);  
    
        }

        private void btTiskEAN_Click(object sender, EventArgs e)
        {

            bool jakTisk = false;
            bool hlasitChybu = true;
            string name = "";
            string num = "";
            string nameZdroj = "";
            string numZdroj = "";
            string namePrvniRadek = ""; // 2A_MM_PPP_DDD
            string nameDruhyRadek = ""; // Vejsada

            //1A Michlova
            //05019001

            // 2A_MM_PPP_DDD
            // Vejsada
            // 0PPPDDD

            numZdroj = txtTextEAN.Text.Trim();
            nameZdroj = txtTextDozimetruEAN.Text.Trim();

            if ((numZdroj.Length == 0) && (hlasitChybu))
                MessageBox.Show("Pole 'Číslo dozimetru' musí být zadáno.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (numZdroj.Length == 0)
            {
                Globalni.Nastroje.LogMessage("Tisk PrintEAN, Pole 'Číslo dozimetru' musí být zadáno, numZdroj:" + numZdroj.ToString(), false, "Error", formRaz);
                return;
            }

            if ((nameZdroj.Length < 4) && (hlasitChybu))
                MessageBox.Show("Pole 'Text dozimetru' musí být minimálně 4 znaky.", Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            // pokud je mensi, doplnim na 4 pozice - je mozne tisknout jen cislo dozimetru beze jmena
            if (nameZdroj.Length < 4)
            {
                Globalni.Nastroje.LogMessage("Tisk PrintEAN, Pole 'Text dozimetru' musí být minimálně 4 znaky., numZdroj:" + nameZdroj.ToString(), false, "Warning", formRaz);
                nameZdroj = nameZdroj.PadLeft(4);
            }

            //namePrvniRadek = nameZdroj.Substring(0, 2) + '_' + numZdroj.Substring(0, 2) + '_' + numZdroj.Substring(2, 3) + '_' + numZdroj.Substring(5, 3);   // 2A_MM_PPP_DDD
            //nameDruhyRadek = nameZdroj.Substring(3, nameZdroj.Length - 3);

            //namePrvniRadek = namePrvniRadek.Replace(" ", "");
            //nameDruhyRadek = nameDruhyRadek.Replace(" ", "");

            //num = '0' + numZdroj.Substring(2, 6);
            //name = namePrvniRadek + " " + nameDruhyRadek;
            num = numZdroj;
            name = nameZdroj;

            if (name.Length > 30)
                name = name.Substring(0, 30);

            //PrintCode39(num, num.Length, name, name.Length);
            // tenhle se pouiva: PrintEAN8(num, num.Length, name, name.Length);


            

            Item itm = (Item)cbEAN.SelectedItem;

            lblStatus.Text = "Tisk "+itm.Name+" ...";
            Globalni.Nastroje.LogMessage("Tisk Print"+itm.Name+" num:" + num.ToString() + ", name: " + name.ToString(), false, "Information", formRaz);

            switch (itm.Value)
            {
                case 1:
                    jakTisk = PrintEAN8(num, num.Length, name, name.Length);
                    break;
                case 2:
                    jakTisk = PrintEAN13(num, num.Length, name, name.Length);
                    break;
                case 3:
                    jakTisk = PrintCode39(num, num.Length, name, name.Length);
                    break;
                default:
                    return ;
            }

            if (jakTisk == true)
            {
                lblStatus.Text = "Tisk " + itm.Name + " ok";
            }
            else
            {
                lblStatus.Text = "Chyba tisku "+ itm.Name ;
                toolStripStatusLabel.Text = "Chyba tisku"+ itm.Name;
                Globalni.Nastroje.LogMessage("Tisk Print" + itm.Name + " num:" + num.ToString() + ", name: " + name.ToString(), false, "Error", formRaz);
                if (hlasitChybu)
                    MessageBox.Show("Chyba tisku "+ itm.Name, Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }



    }
}
