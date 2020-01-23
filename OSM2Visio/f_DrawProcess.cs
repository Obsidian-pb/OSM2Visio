using System;
using System.Windows.Forms;

namespace OSM2Visio
{
    public partial class f_DrawProcess : Form
    {
        public f_DrawProcess()
        {
            InitializeComponent();
        }

        private void B_OK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
