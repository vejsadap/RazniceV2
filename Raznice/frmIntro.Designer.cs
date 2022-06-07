namespace Raznice
{
    partial class frmIntro
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmIntro));
            this.ImgIntroBox = new System.Windows.Forms.PictureBox();
            this.lblInicializace = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.ImgIntroBox)).BeginInit();
            this.SuspendLayout();
            // 
            // ImgIntroBox
            // 
            this.ImgIntroBox.BackColor = System.Drawing.Color.Transparent;
            this.ImgIntroBox.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ImgIntroBox.Cursor = System.Windows.Forms.Cursors.AppStarting;
            this.ImgIntroBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ImgIntroBox.Image = ((System.Drawing.Image)(resources.GetObject("ImgIntroBox.Image")));
            this.ImgIntroBox.Location = new System.Drawing.Point(0, 0);
            this.ImgIntroBox.Name = "ImgIntroBox";
            this.ImgIntroBox.Size = new System.Drawing.Size(300, 214);
            this.ImgIntroBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.ImgIntroBox.TabIndex = 0;
            this.ImgIntroBox.TabStop = false;
            // 
            // lblInicializace
            // 
            this.lblInicializace.BackColor = System.Drawing.Color.Transparent;
            this.lblInicializace.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblInicializace.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.lblInicializace.Location = new System.Drawing.Point(0, 0);
            this.lblInicializace.Name = "lblInicializace";
            this.lblInicializace.Size = new System.Drawing.Size(300, 214);
            this.lblInicializace.TabIndex = 1;
            this.lblInicializace.Text = "Inicializace ...";
            this.lblInicializace.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmIntro
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(300, 214);
            this.Controls.Add(this.ImgIntroBox);
            this.Controls.Add(this.lblInicializace);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmIntro";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmIntro";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.ImgIntroBox)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox ImgIntroBox;
        private System.Windows.Forms.Label lblInicializace;
    }
}