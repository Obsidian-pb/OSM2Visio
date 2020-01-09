using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OSM2Visio.Code
{
    class c_ToolBars
    {
        //private Office.CommandBar MenuBar;
        //private Office.CommandBarPopup newMenuBar;
        //private Office.CommandBarButton ButtonOne;
        //private String menuTag; // As  = "AUniqueName"
        






        /// <summary>Метод создает новую панель инструментов
        /// и отдает команды на добавление двух кнопок на него</summary>
		/// <param name="theApplication">Объект приложения</param>
		/// <param name="commandBarName">Имя создаваемой панели инструментов</param>
        public void CreateCommandBar(
            Microsoft.Office.Interop.Visio.Application theApplication,
            string commandBarName)
        {

            const int FACE_ID = 2934;

            Microsoft.Office.Core.CommandBars theCommandBars;
            Microsoft.Office.Core.CommandBar addedCommandBar;
            Microsoft.Office.Core.CommandBarButton addedCommandBarButton;
            Microsoft.Office.Core.CommandBarButton addedCommandBarButtonFireClick;
            CommandBarEventHandler eventHandler;

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
                addedCommandBar.Context = Convert.ToString((short)Microsoft.Office.Interop.Visio.VisUIObjSets.visUIObjSetDrawing, System.Globalization.CultureInfo.InvariantCulture) + "*";

                // добавляем новую кнопку на  CommandBar.
                addedCommandBarButtonFireClick = (Microsoft.Office.Core.CommandBarButton)
                    addedCommandBar.Controls.Add(Microsoft.Office.Core.
                    MsoControlType.msoControlButton, 1, "", 1, false);

                // Define the button to fire a Click event.
                // This action will be monitored by the event handling class
                // for the Click event.
                // Note: The OnAction property is not used here.
                addedCommandBarButtonFireClick.Caption = "MyButtonClickEvent";
                addedCommandBarButtonFireClick.TooltipText =
                    "Click this button to trigger a Click event";
               
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
                
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
                System.Diagnostics.Debug.WriteLine(err.Message);
                throw;
            }

        }
    }
}
