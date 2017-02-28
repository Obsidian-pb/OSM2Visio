using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OSM2Visio.Code
{
    class CommandBarEventHandler
    {
        //private f_ImportDataDialog FormSelectDialog;

        public void MyCommandBarButtonClick(Office.CommandBarButton cmdButton, ref bool
cancel)
        {
            try
            {
                ThisAddIn.importDataDialogForm.Show();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }

    }
}
