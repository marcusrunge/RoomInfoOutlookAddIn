using Microsoft.Office.Core;
using ModelLibrary;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using static ModelLibrary.Enums;

// TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

// 1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MainRibbon();
//  }

// 2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
//    zu behandeln, z.B. das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem Menüband-Designer exportiert haben,
//    verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und ändern Sie den Code für die Verwendung mit dem
//    Programmmodell für die Menübanderweiterung (RibbonX).

// 3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.  

// Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.


namespace RoomInfoOutlookAddIn
{
    public interface IMainRibbon : IRibbonExtensibility
    {
    }

    [ComVisible(true)]
    public class MainRibbon : IMainRibbon
    {
        private IRibbonUI ribbon;

        public List<RoomItem> RoomItems { get; private set; }
        public List<AgendaItem> AgendaItems { get; private set; }
        public AgendaItem AgendaItem { get; private set; }

        public MainRibbon()
        {

        }

        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("RoomInfoOutlookAddIn.MainRibbon.xml");
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        #endregion

        #region Hilfsprogramme

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        public void OnTcpPortChange(IRibbonControl control)
        {
            
        }

        public void OnUdpPortChange(IRibbonControl control)
        {

        }

        public void OnRoomsDropDownAction(IRibbonControl control)
        {

        }

        public void OnOccupancyDropDownAction(IRibbonControl control)
        {

        }

        public void OnAgendaButtonAction(IRibbonControl control)
        {

        }

        private void ProcessPackage(Package package, string hostName)
        {
            switch ((PayloadType)package.PayloadType)
            {
                case PayloadType.Occupancy:
                    break;
                case PayloadType.Room:
                    if (RoomItems == null) RoomItems = new List<RoomItem>();
                    var room = JsonConvert.DeserializeObject<Room>(package.Payload.ToString());
                        for (int i = 0; i < RoomItems.Count; i++)
                        {
                            if (RoomItems[i].Room.RoomGuid.Equals(room.RoomGuid))
                            {
                                RoomItems.RemoveAt(i);
                                break;
                            }
                        }                    
                    break;
                case PayloadType.Schedule:
                    AgendaItems = new List<AgendaItem>(JsonConvert.DeserializeObject<AgendaItem[]>(package.Payload.ToString()));
                    break;
                case PayloadType.StandardWeek:
                    break;
                case PayloadType.RequestOccupancy:
                    break;
                case PayloadType.RequestSchedule:
                    break;
                case PayloadType.RequestStandardWeek:
                    break;
                case PayloadType.IotDim:
                    break;
                case PayloadType.AgendaItem:
                    break;
                case PayloadType.AgendaItemId:
                    AgendaItem.Id = (int)Convert.ChangeType(package.Payload, typeof(int));
                    break;
                default:
                    break;
            }
        }
    }
}
