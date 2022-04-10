using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;




namespace Raznice.Globalni
{
	/// <summary>
	/// Tøída obsahuje všeobecné statické metody.
	/// </summary>
	public class Nastroje
	{

		/// <summary>
		/// Privátní konstruktor pro znemožnìní vytváøení instancí tøídy.
		/// </summary>
		private Nastroje()
		{

        
		}

		/// <summary>
		/// Vrací cestu právì spuštìné aplikace
		/// </summary>
		/// <returns></returns>
		public static string DejCestuAplikace()
		{
			string cesta = System.IO.Path.GetDirectoryName( 
				System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Replace("file:\\", "");
			
           //ystem.IO.Path.GetDirectoryName(System.Reflection.Assembly.LoadFile(

			if (!cesta.EndsWith("\\"))
				cesta += "\\";

			return cesta;
		}
        public static void LogMessage(string text, bool eventSystemLog, string typUdalosti)
        {
            LogMessage(text, eventSystemLog, typUdalosti, null);
        }

        public static void LogMessage(string text, bool eventSystemLog, string typUdalosti, Object sender)
        {
            try
            {
                LogToFile(text, typUdalosti);
                LogToListBox(text, sender, typUdalosti);
                if (eventSystemLog)
                    switch (typUdalosti)
                    {
                        case "Information":
                            EventLog.WriteEntry(Raznice.Globalni.Parametry.aplikace, text, EventLogEntryType.Information); //EventLogEntryType.Information;
                            break;
                        case "Error":
                            EventLog.WriteEntry(Raznice.Globalni.Parametry.aplikace, text, EventLogEntryType.Error); //EventLogEntryType.Information;
                            break;
                        case "Warning":
                            EventLog.WriteEntry(Raznice.Globalni.Parametry.aplikace, text, EventLogEntryType.Warning); //EventLogEntryType.Information;
                            break;
                        default:
                            EventLog.WriteEntry(Raznice.Globalni.Parametry.aplikace, text, EventLogEntryType.Information); //EventLogEntryType.Information;
                            break;
                    }
            }
            catch (Exception e)
            {
                Console.Write(e.Message.ToString());
            }
        }


        private static void LogToFile(string text, string typUdalosti)
        {
            string logFile;
            logFile = (System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).ToUpper().Replace("EXE", "log");
            logFile = logFile.Replace("FILE:///", "");
            string logFileBackup = logFile;
            logFileBackup = logFileBackup.Replace("log", "old");

            try
            {
                FileInfo fi = new FileInfo(logFile);
                //Console.Write("LogFile: "+fi.FullName.ToString());

                // pokud je log file vetsi nez 1MB
                if (File.Exists(fi.FullName))
                {
                    if (fi.Length > 1000000)
                    {
                        // udelam z .log zalohu .old, minulou zalohu smazu
                        Console.Write("BackupLogFile: " + logFileBackup.ToString());
                        if (File.Exists(logFileBackup))
                            File.Delete(logFileBackup);
                        fi.CopyTo(logFileBackup);
                        fi.Delete();
                    }
                }
                FileInfo fiLog = new FileInfo(logFile);
                DirectoryInfo di = new DirectoryInfo(fiLog.DirectoryName);
                if (!di.Exists) di.Create();

                File.AppendAllText(fiLog.FullName, System.DateTime.Now.ToString() + ": " + "[" + typUdalosti + "] " + text + "\r\n");
            }
             catch (Exception e)
            {
                Console.Write(e.Message.ToString());
            }
        }
        private static void LogToDB(int threadID, int type, string text)
        {
            try
            {
                //Storm.Log.Add(threadID, type, text);
            }
            catch
            {
            }
        }
        private static void LogToScreen(string text)
        {
            try
            {
                Console.WriteLine(text);
            }
            catch
            {
            }
        }

        internal static void PridejTextDoListBoxu(Control f, string text, string typUdalosti)
        {
            foreach (Control c in f.Controls)
            {

                // MessageBox.Show(c.GetType().ToString());
                if (c.HasChildren)
                {
                    PridejTextDoListBoxu(c, text, typUdalosti);
                }
                else
                    if (c is ListBox)
                    {
                        ListBox lll = (ListBox)c;
                        if (lll.Name == "listBoxLog")
                        {
                            //MessageBox.Show(lll.Name.ToString());
                            lll.Items.Add(System.DateTime.Now.ToString() + ": " + "["+typUdalosti+"] " + text);

                        }
                    }
            }
        }

        private static void LogToListBox(string text, Object sender, string typUdalosti)
        {
            try
            {
                Vlastnosti.logText = text;

                //frmRaznice form = new frmRaznice();

                if (sender != null)
                    PridejTextDoListBoxu((Form)sender, text, typUdalosti);
 
            }
            catch
            {
            }
        }

	}

    

}
