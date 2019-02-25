using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using System.Globalization;
using Unity;
using Microsoft.Office.Core;
using ApplicationServiceLibrary;
using ModelLibrary;

namespace RoomInfoOutlookAddIn
{
    public partial class ThisAddIn
    {
        IUnityContainer _unityContainer;
        IEventService _eventService;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Outlook.MAPIFolder calendar = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);            
            Outlook.Items calendarItems = calendar.Items;
            calendarItems.ItemAdd += CalendarItems_ItemAdd;
            calendarItems.ItemChange += CalendarItems_ItemChange;
            calendarItems.ItemRemove += CalendarItems_ItemRemove;
            _eventService = _unityContainer.Resolve<IEventService>();
            _eventService.AddButtonPressed += ThisAddIn_AddButtonPressed;
        }

        private void ThisAddIn_AddButtonPressed(object sender, RoomItem roomItem)
        {
            
        }

        private void CalendarItems_ItemRemove()
        {
            
        }

        private void CalendarItems_ItemChange(object Item)
        {
            
        }

        private void CalendarItems_ItemAdd(object Item)
        {
            
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _unityContainer = new UnityContainer();
            _unityContainer.RegisterSingleton<IMainRibbon, MainRibbon>();
            _unityContainer.RegisterSingleton<INetworkCommunication, NetworkCommunication>();
            _unityContainer.RegisterSingleton<IEventService, EventService>();
            Outlook.Application outlookApplication = GetHostItem<Outlook.Application>(typeof(Outlook.Application), "Application");
            int languageID = outlookApplication.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(languageID);
            return _unityContainer.Resolve<IMainRibbon>();
        }
    }
}
