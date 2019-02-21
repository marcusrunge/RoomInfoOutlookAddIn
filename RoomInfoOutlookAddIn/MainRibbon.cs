using ApplicationServiceLibrary;
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
        private INetworkCommunication _networkCommunication;

        public List<RoomItem> RoomItems { get; private set; }
        public List<AgendaItem> AgendaItems { get; private set; }
        public AgendaItem AgendaItem { get; private set; }

        public MainRibbon(INetworkCommunication networkCommunication)
        {
            _networkCommunication = networkCommunication;
            _networkCommunication.StartConnectionListener(Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
            _networkCommunication.PayloadReceived += (s, e) => ProcessPackage(JsonConvert.DeserializeObject<Package>(e.Package), e.HostName);
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

        public void OnChange(IRibbonControl control, string text)
        {
            int parsed;
            switch (control.Id)
            {
                case "tcpPort":
                    if (int.TryParse(text, out parsed))
                    {
                        Properties.Settings.Default.TcpPort = text;
                        Properties.Settings.Default.Save();
                    }
                    break;
                case "udpPort":
                    if (int.TryParse(text, out parsed))
                    {
                        Properties.Settings.Default.UdpPort = text;
                        Properties.Settings.Default.Save();
                    }
                    break;
                default:
                    break;
            }
        }

        public string OnGetText(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "tcpPort": return Properties.Settings.Default.TcpPort;
                case "udpPort": return Properties.Settings.Default.UdpPort;
                default: return "";
            }
        }

        public void OnRoomsDropDownAction(IRibbonControl control)
        {

        }

        public void OnOccupancyDropDownAction(IRibbonControl control)
        {

        }

        public void OnAction(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "agendaButton":
                    break;
                case "recycleButton":
                    _networkCommunication.SendPayload("", null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
                    break;
                default:
                    break;
            }
        }

        private void ProcessPackage(Package package, string hostName)
        {
            switch ((PayloadType)package.PayloadType)
            {
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
                case PayloadType.AgendaItemId:
                    AgendaItem.Id = (int)Convert.ChangeType(package.Payload, typeof(int));
                    break;
                default:
                    break;
            }
        }
    }
}
