using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Raznice
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                /*
                Counter c = new Counter(new Random().Next(10));
                c.ThresholdReached += c_ThresholdReached;
                */
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                frmLogin fLogin = new frmLogin();
                DialogResult vyslLogin = fLogin.ShowDialog();
                            
                if (vyslLogin == DialogResult.Abort)
                {
                    fLogin.Close();
                    Vlastnosti.exit = true;

                }
                Application.Run(new frmRaznice());
                
                
            }
            catch (Exception e)
            {
                string chyba = "Source:" + e.Source.ToString() + 
                                ", Message:" + e.Message.ToString() + 
                                ", Stack:"+ e.StackTrace.ToString() +
                                ", TargetSite:" + e.TargetSite.ToString() +
                                ", Data:" + e.Data.ToString();
                Globalni.Nastroje.LogMessage(Globalni.Parametry.aplikace.ToString()+ ": " + chyba, false, "Error");
                MessageBox.Show(chyba, Globalni.Parametry.aplikace.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



        }

        static void c_ThresholdReached(object sender, EventArgs e)
        {
            Console.WriteLine("The threshold was reached.");
            Environment.Exit(0);
        }
    }

}
