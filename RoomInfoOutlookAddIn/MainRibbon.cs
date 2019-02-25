using ApplicationServiceLibrary;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using ModelLibrary;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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

        private List<RoomItem> _roomItems;
        private AgendaItem _agendaItem;
        private int _selectedRoomId;
        private ResourceManager _resourceManager;

        public MainRibbon(INetworkCommunication networkCommunication)
        {
            var cultureInfo = Thread.CurrentThread.CurrentUICulture;
            Properties.Resources.Culture = cultureInfo;
            _resourceManager = Properties.LanguageResources.ResourceManager;
            _selectedRoomId = 0;
            _networkCommunication = networkCommunication;
            _roomItems = new List<RoomItem>();
            _agendaItem = new AgendaItem();
            _networkCommunication.StartConnectionListener(Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
            _networkCommunication.PayloadReceived += async (s, e) =>
            {
                if (e.Package != null) await ProcessPackage(JsonConvert.DeserializeObject<Package>(e.Package), e.HostName);
            };
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

        public string OnGetLabel(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "MainTab": return _resourceManager.GetString("Tab_Label");
                case "recycleButton": return _resourceManager.GetString("RecycleButton_Label");
                default: return "";
            }
        }

        public int GetItemCount(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "roomsDropDown": return _roomItems != null ? _roomItems.Count : 0;
                case "occupancyDropDown": return 6;
                default: return 0;
            }

        }

        public async void OnAction(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "agendaButton":
                    break;
                case "recycleButton":
                    await _networkCommunication.SendPayload("", null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
                    break;
                default:
                    break;
            }
        }

        public async void OnDropDownAction(IRibbonControl control, string itemID, int itemIndex)
        {
            switch (control.Id)
            {
                case "roomsDropDown":
                    _selectedRoomId = itemIndex;
                    ribbon.Invalidate();
                    break;
                case "occupancyDropDown":
                    if (_roomItems != null && _roomItems.Count >= _selectedRoomId + 1)
                    {
                        _roomItems[_selectedRoomId].Room.Occupancy = itemIndex;
                        var package = new Package() { PayloadType = (int)PayloadType.Occupancy, Payload = itemIndex };
                        await _networkCommunication.SendPayload(JsonConvert.SerializeObject(package), _roomItems[_selectedRoomId].HostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
                    }
                    break;
                default:
                    break;
            }
        }

        public string GetItemLabel(IRibbonControl control, int index)
        {
            switch (control.Id)
            {
                case "roomsDropDown":
                    var roomNumber = _roomItems[index].Room.RoomNumber;
                    var roomName = _roomItems[index].Room.RoomName;
                    return !(string.IsNullOrEmpty(roomNumber) || string.IsNullOrWhiteSpace(roomNumber))
                        ? !(string.IsNullOrEmpty(roomName) || string.IsNullOrWhiteSpace(roomName)) ? roomName + " " + roomNumber : roomNumber
                        : !(string.IsNullOrEmpty(roomName) || string.IsNullOrWhiteSpace(roomName)) ? roomName : "";
                case "occupancyDropDown":
                    switch (index)
                    {
                        case 0: return _resourceManager.GetString("DropDown_ItemLabel_Occupancy_Free");
                        case 1: return _resourceManager.GetString("DropDown_ItemLabel_Occupancy_Present");
                        case 2: return _resourceManager.GetString("DropDown_ItemLabel_Occupancy_Absent");
                        case 3: return _resourceManager.GetString("DropDown_ItemLabel_Occupancy_Busy");
                        case 4: return _resourceManager.GetString("DropDown_ItemLabel_Occupancy_Occupied");
                        case 5: return _resourceManager.GetString("DropDown_ItemLabel_Occupancy_Locked");
                        default: return "";
                    }
                default: return "";
            }
        }

        public string GetItemID(IRibbonControl control, int index) => _roomItems[index].Room.RoomGuid;

        public int GetSelectedItemIndex(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "roomsDropDown": return _selectedRoomId;
                case "occupancyDropDown": return _roomItems != null && _roomItems.Count >= _selectedRoomId + 1 ? _roomItems[_selectedRoomId].Room.Occupancy : 0;
                default: return 0;
            }
        }

        private async Task ProcessPackage(Package package, string hostName)
        {
            switch ((PayloadType)package.PayloadType)
            {
                case PayloadType.Room:
                    if (_roomItems == null) _roomItems = new List<RoomItem>();
                    var room = JsonConvert.DeserializeObject<Room>(package.Payload.ToString());
                    for (int i = 0; i < _roomItems.Count; i++)
                    {
                        if (_roomItems[i].Room.RoomGuid.Equals(room.RoomGuid))
                        {
                            _roomItems.RemoveAt(i);
                            break;
                        }
                    }
                    _roomItems.Add(new RoomItem() { HostName = hostName, Room = room });
                    package = new Package() { PayloadType = (int)PayloadType.RequestSchedule };
                    await _networkCommunication.SendPayload(JsonConvert.SerializeObject(package), hostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
                    ribbon.Invalidate();
                    break;
                case PayloadType.Schedule:
                    var agendaItems = new List<AgendaItem>(JsonConvert.DeserializeObject<AgendaItem[]>(package.Payload.ToString()));
                    for (int i = 0; i < _roomItems.Count; i++)
                    {
                        if (_roomItems[i].HostName == hostName)
                        {
                            _roomItems[i].AgendaItems = agendaItems;
                            break;
                        }
                    }
                    break;
                case PayloadType.StandardWeek:
                    break;
                case PayloadType.AgendaItemId:
                    _agendaItem.Id = (int)Convert.ChangeType(package.Payload, typeof(int));
                    break;
                default:
                    break;
            }
        }
    }
}
