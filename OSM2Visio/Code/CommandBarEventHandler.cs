using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OSM2Visio.Code
{
    class CommandBarEventHandler
    {
        private f_ImportDataDialog FormSelectDialog;

        public void MyCommandBarButtonClick(Office.CommandBarButton cmdButton, ref bool cancel)
        {
            //MessageBox.Show("You Did it!!!");
            FormSelectDialog = new f_ImportDataDialog();
            FormSelectDialog.Show();
            //FormSelectDialog.;
        }
    }
}
