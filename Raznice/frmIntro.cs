using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Raznice
{
    public partial class frmIntro : Form
    {
        public frmIntro()
        {
            InitializeComponent();
            lblInicializace.Parent = ImgIntroBox; // transparentni label
            lblInicializace.Text = "Inicializace spojení ...";
            this.TransparencyKey = (BackColor); // transparentni form
        }
    }
}
