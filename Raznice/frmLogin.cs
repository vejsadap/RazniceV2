using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Raznice
{
    public partial class frmLogin : Form
    {
        public frmLogin()
        {
            InitializeComponent();
        }

        private void CmdVysledek_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.DialogResult vyslOK = System.Windows.Forms.DialogResult.OK;
            System.Windows.Forms.DialogResult vyslNo = System.Windows.Forms.DialogResult.No;

            this.DialogResult = vyslNo;
            Vlastnosti.allowEdit = false;

            if ((this.cmbUroven.Text == "Administrator") && (this.txtHeslo.Text.Trim() == "hasanasan"))
            {               
                this.DialogResult = vyslOK;
                Vlastnosti.allowEdit = true;
            }
            if (this.cmbUroven.Text != "Administrator")
            {
                this.DialogResult = vyslOK;
                Vlastnosti.allowEdit = false;
            }
            this.Close();
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            this.cmbUroven.Items.Add("Uživatel");
            this.cmbUroven.Items.Add("Administrator");

            this.cmbUroven.Text = "Uživatel";
        }

        private void cmbUroven_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cmbUroven.Text == "Uživatel")
            {
                this.txtHeslo.Enabled = false;
                this.lblHeslo.Enabled = false;
                this.cmdVysledek.Focus();
            }
            else
            {
                this.txtHeslo.Enabled = true;
                this.lblHeslo.Enabled = true;
                this.txtHeslo.Focus();
            }
        }


    }
}
