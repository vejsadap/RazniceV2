namespace Raznice
{
    partial class frmLogin
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmLogin));
            this.cmdVysledek = new System.Windows.Forms.Button();
            this.cmdNo = new System.Windows.Forms.Button();
            this.cmbUroven = new System.Windows.Forms.ComboBox();
            this.lblUroven = new System.Windows.Forms.Label();
            this.lblHeslo = new System.Windows.Forms.Label();
            this.txtHeslo = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmdVysledek
            // 
            this.cmdVysledek.Location = new System.Drawing.Point(20, 91);
            this.cmdVysledek.Name = "cmdVysledek";
            this.cmdVysledek.Size = new System.Drawing.Size(75, 23);
            this.cmdVysledek.TabIndex = 0;
            this.cmdVysledek.Text = "Ok";
            this.cmdVysledek.UseVisualStyleBackColor = true;
            this.cmdVysledek.Click += new System.EventHandler(this.CmdVysledek_Click);
            // 
            // cmdNo
            // 
            this.cmdNo.DialogResult = System.Windows.Forms.DialogResult.Abort;
            this.cmdNo.Location = new System.Drawing.Point(126, 91);
            this.cmdNo.Name = "cmdNo";
            this.cmdNo.Size = new System.Drawing.Size(75, 23);
            this.cmdNo.TabIndex = 5;
            this.cmdNo.Text = "Storno";
            this.cmdNo.UseVisualStyleBackColor = true;
            // 
            // cmbUroven
            // 
            this.cmbUroven.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbUroven.FormattingEnabled = true;
            this.cmbUroven.Location = new System.Drawing.Point(79, 19);
            this.cmbUroven.Name = "cmbUroven";
            this.cmbUroven.Size = new System.Drawing.Size(121, 21);
            this.cmbUroven.TabIndex = 1;
            this.cmbUroven.SelectedIndexChanged += new System.EventHandler(this.cmbUroven_SelectedIndexChanged);
            // 
            // lblUroven
            // 
            this.lblUroven.AutoSize = true;
            this.lblUroven.Location = new System.Drawing.Point(25, 22);
            this.lblUroven.Name = "lblUroven";
            this.lblUroven.Size = new System.Drawing.Size(42, 13);
            this.lblUroven.TabIndex = 2;
            this.lblUroven.Text = "Úroveň";
            // 
            // lblHeslo
            // 
            this.lblHeslo.AutoSize = true;
            this.lblHeslo.Location = new System.Drawing.Point(25, 48);
            this.lblHeslo.Name = "lblHeslo";
            this.lblHeslo.Size = new System.Drawing.Size(34, 13);
            this.lblHeslo.TabIndex = 3;
            this.lblHeslo.Text = "Heslo";
            // 
            // txtHeslo
            // 
            this.txtHeslo.Location = new System.Drawing.Point(79, 45);
            this.txtHeslo.Name = "txtHeslo";
            this.txtHeslo.Size = new System.Drawing.Size(121, 20);
            this.txtHeslo.TabIndex = 4;
            this.txtHeslo.UseSystemPasswordChar = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtHeslo);
            this.groupBox1.Controls.Add(this.lblHeslo);
            this.groupBox1.Controls.Add(this.lblUroven);
            this.groupBox1.Controls.Add(this.cmbUroven);
            this.groupBox1.Location = new System.Drawing.Point(4, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(224, 83);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Přihlašovací údaje";
            // 
            // frmLogin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(233, 116);
            this.ControlBox = false;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.cmdNo);
            this.Controls.Add(this.cmdVysledek);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmLogin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Přihlášení";
            this.Load += new System.EventHandler(this.frmLogin_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cmdVysledek;
        private System.Windows.Forms.Button cmdNo;
        private System.Windows.Forms.ComboBox cmbUroven;
        private System.Windows.Forms.Label lblUroven;
        private System.Windows.Forms.Label lblHeslo;
        private System.Windows.Forms.TextBox txtHeslo;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}