using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
using System.Windows.Forms;

namespace OSM2Visio
{
    public partial class f_DrawProcess : Form
    {
        private bool _isGoon;

        public f_DrawProcess()
        {
            _isGoon = true;
            InitializeComponent();
        }

        public bool IsGoon
        {
            get { return _isGoon; }
            set { _isGoon = value; }
        }

        private void B_OK_Click(object sender, EventArgs e)
        {
            this.Hide();
        }


        void f_DrawProcess_Closed(object sender, EventArgs e)
        {
            _isGoon = false;
            return;
        }

    }
}
