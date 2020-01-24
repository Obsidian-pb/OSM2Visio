using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using OSM2Visio.Properties;
using Visio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace OSM2Visio
{
    public partial class ThisAddIn
    {
        private readonly AddinUI AddinUI = new AddinUI();

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return AddinUI;
        }

        /// <summary>
        /// A simple command
        /// </summary>
        public void OsmImport()
        {
            var FormSelectDialog = new f_ImportDataDialog();
            FormSelectDialog.Show();
        }

        /// <summary>
        /// Callback called by the UI manager when user clicks a button
        /// Should do something meaningful when corresponding action is called.
        /// </summary>
        public void OnCommand(string commandId)
        {
            switch (commandId)
            {
                case "OsmImport":
                    OsmImport();
                    return;
            }
        }

        /// <summary>
        /// Callback called by UI manager.
        /// Should return if corresponding command should be enabled in the user interface.
        /// By default, all commands are enabled.
        /// </summary>
        public bool IsCommandEnabled(string commandId)
        {
            switch (commandId)
            {
                case "OsmImport":    // make command1 always enabled
                    return true;

                default:
                    return true;
            }
        }

        public bool IsCommandChecked(string command)
        {
            return false;
        }

        public string GetCommandText(string command, string suffix)
        {
            return Resources.ResourceManager.GetString($"{command}_{suffix}", Resources.Culture);
        }

        public Bitmap GetCommandBitmap(string id)
        {
            return (Bitmap)Resources.ResourceManager.GetObject(id, Resources.Culture);
        }

        internal void UpdateUI()
        {
            AddinUI.UpdateCommandBars();
            AddinUI.UpdateRibbon();
        }

        private void Application_SelectionChanged(Visio.Window window)
        {
            UpdateUI();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var version = int.Parse(Application.Version, NumberStyles.AllowDecimalPoint);
            if (version < 14)
                AddinUI.StartupCommandBars("OSM2Visio", new[] { "OsmImport" });
            Application.SelectionChanged += Application_SelectionChanged;

        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            AddinUI.ShutdownCommandBars();
            Application.SelectionChanged -= Application_SelectionChanged;

        }


        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            // TODO: add proper language switching
            var lcid = Application.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
            var culture = (lcid != 1033)
                ? new CultureInfo(lcid)
                : CultureInfo.CurrentCulture;

            var languages = new[] { "ru" };
            if (languages.Any(language => language == culture.TwoLetterISOLanguageName))
            {
                Resources.Culture = culture;
                System.Threading.Thread.CurrentThread.CurrentUICulture = culture;
                UpdateUI();
            }

            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

    }
}
