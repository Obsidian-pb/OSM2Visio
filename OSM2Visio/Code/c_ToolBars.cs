using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OSM2Visio.Code
{
    class c_ToolBars
    {
        static CommandBarEventHandler eventHandler;

        Microsoft.Office.Interop.Visio.Application visioApplication;
        public Microsoft.Office.Core.CommandBars theCommandBars;
        public Microsoft.Office.Core.CommandBar addedCommandBar;
        public Microsoft.Office.Core.CommandBarButton addedCommandBarButtonFireClick;

        /// <summary>Метод создает новую панель инструментов
        /// и отдает команды на добавление двух кнопок на него</summary>
		/// <param name="theApplication">Объект приложения</param>
		/// <param name="commandBarName">Имя создаваемой панели инструментов</param>
        public void CreateCommandBar(
            Microsoft.Office.Interop.Visio.Application theApplication,
            string commandBarName)
        {
            const int FACE_ID = 2934;

            //Сохраняем ссылку на приложение Visio
            visioApplication = theApplication;

            //Подписываемся на событие "Открытие документа" (при открытии нового документа проверяем, не является ли он трафаретом "Водоснабжение" или "План на местности")
            theApplication.Documents.DocumentOpened += new Visio.EDocuments_DocumentOpenedEventHandler(Documents_DocumentOpened);
            //Подписываемся на событие "Закрытие документа" (при закрытии существующего документа проверяем, не является ли он трафаретом "Водоснабжение" или "План на местности")
            theApplication.Documents.BeforeDocumentClose += new Visio.EDocuments_BeforeDocumentCloseEventHandler(Documents_BeforeDocumentClose);

            if (commandBarName == null || theApplication == null)
            {
                return;
            }

            try
            {

                // Уточняем указан ли атрибут имени панели инструментов.
                if (commandBarName.Length == 0)
                {
                    throw new System.ArgumentNullException("commandBarName",
                        "Zero length string input.");
                }

                eventHandler = new CommandBarEventHandler();


                // Добавляем новую панель инструментов.
                theCommandBars = (Microsoft.Office.Core.CommandBars) theApplication.CommandBars;
                addedCommandBar = theCommandBars.Add(commandBarName,
                    Microsoft.Office.Core.MsoBarPosition.msoBarTop,
                    false, true);

                // Запрещаем пользователю самостоятельно модифицировать панель.
                addedCommandBar.Protection = Microsoft.Office.Core.MsoBarProtection.msoBarNoCustomize;

                // Отображаем the command bar.
                //addedCommandBar.Context = Convert.ToString((short)Microsoft.Office.Interop.Visio.VisUIObjSets.visUIObjSetDrawing, System.Globalization.CultureInfo.InvariantCulture) + "*";
                addedCommandBar.Visible = true;

                // добавляем новую кнопку на  CommandBar.
                addedCommandBarButtonFireClick = (Microsoft.Office.Core.CommandBarButton)
                    addedCommandBar.Controls.Add(Microsoft.Office.Core.
                    MsoControlType.msoControlButton, 1, "", 1, false);

                // Define the button to fire a Click event.
                // This action will be monitored by the event handling class
                // for the Click event.
                // Note: The OnAction property is not used here.
                addedCommandBarButtonFireClick.Caption = "Импорт карты OSM";
                addedCommandBarButtonFireClick.TooltipText =
                    "Импорт карты OSM";
               
                // Put the button in a group bar.
                addedCommandBarButtonFireClick.BeginGroup = true;
                
                // Use the Tag property for context switching
                // and for use with the FindControl method.
                addedCommandBarButtonFireClick.Tag = "clickEvent";
                
                // Get an internal icon for the button.
                addedCommandBarButtonFireClick.FaceId = FACE_ID;
                
                // Set a reference to the CommandBar button in the
                // event handling class.
                addedCommandBarButtonFireClick.Click += new Microsoft.Office.Core.
                    _CommandBarButtonEvents_ClickEventHandler
                    (eventHandler.MyCommandBarButtonClick);

                //addedCommandBarButtonFireClick.Enabled = false;
                addedCommandBarButtonFireClick.Enabled = IsNeededTrafaretsExists();
                
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                System.Diagnostics.Debug.WriteLine(err.Message);
                throw;
            }

        }

        #region Проки проверки наличия необходимых открытых трафаретов
        void Documents_DocumentOpened(Visio.Document Doc)
        {
            try
            {
                Microsoft.Office.Core.CommandBars commandBars;
                Microsoft.Office.Core.CommandBar commandBar;
                Microsoft.Office.Core.CommandBarButton button;
                commandBars = visioApplication.CommandBars;
                commandBar = commandBars["OSM Import"];
                button = (Microsoft.Office.Core.CommandBarButton)commandBar.Controls["Импорт карты OSM"];

                button.Enabled = IsNeededTrafaretsExists();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }

        void Documents_BeforeDocumentClose(Visio.Document Doc)
        {
            try
            {
                Microsoft.Office.Core.CommandBars commandBars;
                Microsoft.Office.Core.CommandBar commandBar;
                Microsoft.Office.Core.CommandBarButton button;
                commandBars = visioApplication.CommandBars;
                commandBar = commandBars["OSM Import"];
                button = (Microsoft.Office.Core.CommandBarButton)commandBar.Controls["Импорт карты OSM"];

                if (IsNeededTrafaretsExists())
                    button.Enabled = true;
                else
                    button.Enabled = false;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }
        }

        private bool IsNeededTrafaretsExists()
        {
            bool stencilWaterSourceExists = false;
            bool stencilPlacePlanExists = false;

            foreach (Visio.Document doc in visioApplication.Documents)
            {
                if (doc.Name == "Водоснабжение.vss") stencilWaterSourceExists = true;
                if (doc.Name == "План на местности.vss") stencilPlacePlanExists = true;
            }
            //MessageBox.Show(stencilWaterSourceExists.ToString() + " - " + stencilPlacePlanExists.ToString());
            //Проверяем оба ли трафарета имеются в документе
            if (stencilWaterSourceExists && stencilPlacePlanExists)
            {
                return true;
            }
            return false;
        }
        #endregion Проки проверки наличия необходимых открытых трафаретов
    }
}
