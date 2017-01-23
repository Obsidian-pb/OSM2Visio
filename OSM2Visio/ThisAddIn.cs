using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using OSM2Visio.Code;
using System.Windows.Forms;

namespace OSM2Visio
{
    public partial class ThisAddIn
    {
        private Visio.Application VisApp;
        private c_ToolBars TBO;
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Получаем ссылку на текущее приложение
            Visio.Application VisApp = Globals.ThisAddIn.Application;
            //Получаем ссылку на набор панелей инструментов
            TBO = new c_ToolBars();
            //Создаем новую панель инструментов
            TBO.CreateCommandBar(VisApp, "OSM Import");

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //
            VisApp = null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
