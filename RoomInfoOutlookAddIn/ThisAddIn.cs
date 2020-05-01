using ApplicationServiceLibrary;
using Microsoft.Office.Core;
using ModelLibrary;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Unity;
using static ModelLibrary.Enums;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RoomInfoOutlookAddIn
{
    public partial class ThisAddIn
    {
        IUnityContainer _unityContainer;
        IEventService _eventService;
        INetworkCommunication _networkCommunication;
        IList<RoomItem> _roomItems;
        Outlook.AppointmentItem _appointmentItem/*, _openAppointmentItem*/;
        Outlook.MAPIFolder _roomInfocalendar;
        Package _discoveryPackage;
        bool _isSyncInProgress;
        //Outlook.Explorer currentExplorer = null;

        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _roomInfocalendar = CreateCustomCalendarIfNotExists("RoomInfoCalendar");
            _roomInfocalendar.Items.ItemAdd += CalendarItems_ItemAdd;
            _roomInfocalendar.Items.ItemChange += CalendarItems_ItemChange;
            _roomInfocalendar.Items.ItemRemove += CalendarItems_ItemRemove;
            _eventService = _unityContainer.Resolve<IEventService>();
            _roomItems = _unityContainer.Resolve<IList<RoomItem>>();
            _eventService.AddButtonPressed += ThisAddIn_AddButtonPressed;
            _eventService.SyncButtonPressed += _eventService_SyncButtonPressed;
            _networkCommunication.PayloadReceived += async (s, payload) =>
            {
                if (payload.Package != null) await ProcessPackage(JsonConvert.DeserializeObject<Package>(payload.Package), payload.HostName);
            };
            await _networkCommunication.SendPayload(JsonConvert.SerializeObject(_discoveryPackage), null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
            _eventService.ScheduleReceived += (s, roomItem) =>
            {
                AddItemsToCalendar(roomItem.AgendaItems, roomItem.Room);
                _isSyncInProgress = false;
            };
            //currentExplorer = Application.ActiveExplorer();
            //currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);
        }

        //private void CurrentExplorer_Event()
        //{
        //    Outlook.MAPIFolder selectedFolder = Application.ActiveExplorer().CurrentFolder;            
        //    try
        //    {
        //        if (Application.ActiveExplorer().Selection.Count > 0)
        //        {
        //            object selObject = Application.ActiveExplorer().Selection[1];
        //            if (selObject is Outlook.AppointmentItem)
        //            {
        //                if (_openAppointmentItem != null) _openAppointmentItem.Close(Outlook.OlInspectorClose.olDiscard);
        //                _openAppointmentItem = (selObject as Outlook.AppointmentItem);
        //                _openAppointmentItem.Open += (ref bool c) => _eventService.OnOutlookAppointmentItemOpen(_openAppointmentItem);
        //            }                    
        //        }
        //    }
        //    catch (Exception)
        //    {

        //    }
        //}

        private async void CalendarItems_ItemRemove()
        {
            if (_isSyncInProgress) return;
            try
            {
                foreach (var roomItem in _roomItems)
                {
                    for (int i = 0; i < roomItem.AgendaItems.Count; i++)
                    {
                        int id = roomItem.AgendaItems[i].Id;
                        if (id > 0)
                        {
                            string guid = roomItem.Room.RoomGuid;
                            Outlook.AppointmentItem appointmentItem = null;
                            foreach (Outlook.AppointmentItem item in _roomInfocalendar.Items)
                            {
                                if (item.Resources == null) break;
                                string itemGuid = item.UserProperties.Find("RoomGuid").Value;
                                int itemId = int.Parse(item.UserProperties.Find("RemoteDbEntityId").Value);
                                if (itemGuid.Equals(guid) && itemId == id)
                                {
                                    appointmentItem = item;
                                    break;
                                }
                            }
                            if (appointmentItem == null)
                            {
                                roomItem.AgendaItems.Remove(roomItem.AgendaItems[i]);
                                var agendaItem = new AgendaItem() { Id = id, IsDeleted = true };
                                var agendaItemPackage = new Package() { PayloadType = (int)PayloadType.AgendaItem, Payload = agendaItem };
                                await _networkCommunication.SendPayload(JsonConvert.SerializeObject(agendaItemPackage), roomItem.HostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
                            }
                        }
                    }
                    var schedulePackage = new Package() { PayloadType = (int)PayloadType.RequestSchedule };
                    await _networkCommunication.SendPayload(JsonConvert.SerializeObject(schedulePackage), roomItem.HostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
                }
            }
            catch (Exception)
            {

            }
        }

        private async void _eventService_SyncButtonPressed(object sender, RoomItem roomItem)
        {
            _isSyncInProgress = true;
            ClearCalendar(roomItem);
            var schedulePackage = new Package() { PayloadType = (int)PayloadType.RequestSchedule };
            await _networkCommunication.SendPayload(JsonConvert.SerializeObject(schedulePackage), roomItem.HostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
        }

        private void ThisAddIn_AddButtonPressed(object sender, RoomItem roomItem)
        {
            try
            {
                var roomNumber = roomItem.Room.RoomNumber;
                var roomName = roomItem.Room.RoomName;
                Outlook.AppointmentItem appointmentItem = _roomInfocalendar.Items.Add(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                appointmentItem.Location = !(string.IsNullOrEmpty(roomNumber) || string.IsNullOrWhiteSpace(roomNumber))
                    ? !(string.IsNullOrEmpty(roomName) || string.IsNullOrWhiteSpace(roomName)) ? roomName + " " + roomNumber : roomNumber
                    : !(string.IsNullOrEmpty(roomName) || string.IsNullOrWhiteSpace(roomName)) ? roomName : "";
                var remoteDbEntityId = appointmentItem.UserProperties.Add("RemoteDbEntityId", Outlook.OlUserPropertyType.olInteger);
                var remoteDbEntityTimeStamp = appointmentItem.UserProperties.Add("RemoteDbEntityTimeStamp", Outlook.OlUserPropertyType.olText);
                var roomGuid = appointmentItem.UserProperties.Add("RoomGuid", Outlook.OlUserPropertyType.olText);
                remoteDbEntityId.Value = 0;
                remoteDbEntityTimeStamp.Value = DateTimeOffset.Now.ToUnixTimeMilliseconds().ToString();
                roomGuid.Value = roomItem.Room.RoomGuid;
                appointmentItem.BusyStatus = Outlook.OlBusyStatus.olBusy;
                appointmentItem.Display();
            }
            catch (Exception)
            {

            }
        }

        private async void CalendarItems_ItemChange(object Item)
        {
            if (_isSyncInProgress) return;
            _appointmentItem = Item as Outlook.AppointmentItem;
            await TransmitAgendaItem(_appointmentItem);
        }

        private async void CalendarItems_ItemAdd(object Item)
        {
            if (_isSyncInProgress) return;
            try
            {
                _appointmentItem = Item as Outlook.AppointmentItem;
                await TransmitAgendaItem(_appointmentItem);
                string hostName = GetHostName(_roomItems, _appointmentItem);
                var schedulePackage = new Package() { PayloadType = (int)PayloadType.RequestSchedule };
                await _networkCommunication.SendPayload(JsonConvert.SerializeObject(schedulePackage), hostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
            }
            catch (Exception)
            {

            }
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
            _isSyncInProgress = false;
            _discoveryPackage = new Package() { PayloadType = (int)PayloadType.Discovery };
            _unityContainer = new UnityContainer();
            _unityContainer.RegisterSingleton<IMainRibbon, MainRibbon>();
            _unityContainer.RegisterSingleton<INetworkCommunication, NetworkCommunication>();
            _unityContainer.RegisterSingleton<IEventService, EventService>();
            _unityContainer.RegisterSingleton<IList<RoomItem>, List<RoomItem>>();
            //UnityServiceLocator unityServiceLocator = new UnityServiceLocator(_unityContainer);
            //ServiceLocator.SetLocatorProvider(() => unityServiceLocator);
            _networkCommunication = _unityContainer.Resolve<INetworkCommunication>();
            Task.Run(async () => await _networkCommunication.StartConnectionListener(Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl));
            Task.Run(async () => await _networkCommunication.StartConnectionListener(Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram));
            Outlook.Application outlookApplication = GetHostItem<Outlook.Application>(typeof(Outlook.Application), "Application");
            int languageID = outlookApplication.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(languageID);
            return _unityContainer.Resolve<IMainRibbon>();
        }

        private async Task TransmitAgendaItem(Outlook.AppointmentItem appointmentItem)
        {
            string hostName = GetHostName(_roomItems, appointmentItem);
            var updatedAgendaItem = new AgendaItem()
            {
                Id = appointmentItem.UserProperties.Find("RemoteDbEntityId").Value,
                Description = appointmentItem.Body,
                End = appointmentItem.End,
                IsAllDayEvent = appointmentItem.AllDayEvent,
                Start = appointmentItem.Start,
                Title = appointmentItem.Subject,
                TimeStamp = long.Parse(appointmentItem.UserProperties.Find("RemoteDbEntityTimeStamp").Value),
                Occupancy = appointmentItem.UserProperties.Find("Occupancy").Value
            };
            string guid = appointmentItem.UserProperties.Find("RoomGuid").Value;
            var agendaItems = _roomItems.Where(x => x.Room.RoomGuid.Equals(guid)).Select(x => x.AgendaItems).FirstOrDefault();
            if (agendaItems != null)
            {
                var agendaItem = agendaItems.Where(x => x.Id == updatedAgendaItem.Id).Select(x => x).FirstOrDefault();
                if (agendaItem != null &&
                    agendaItem.IsAllDayEvent == updatedAgendaItem.IsAllDayEvent &&
                    agendaItem.Occupancy == updatedAgendaItem.Occupancy &&
                    agendaItem.Start == updatedAgendaItem.Start &&
                    agendaItem.TimeStamp == updatedAgendaItem.TimeStamp &&
                    agendaItem.Title == updatedAgendaItem.Title &&
                    agendaItem.Description == updatedAgendaItem.Description &&
                    agendaItem.End == updatedAgendaItem.End)
                    return;
                var unixTimeMilliseconds = DateTimeOffset.Now.ToUnixTimeMilliseconds();
                updatedAgendaItem.TimeStamp = unixTimeMilliseconds;
                appointmentItem.UserProperties.Find("RemoteDbEntityTimeStamp").Value = unixTimeMilliseconds.ToString();
                var package = new Package() { PayloadType = (int)PayloadType.AgendaItem, Payload = updatedAgendaItem };
                await _networkCommunication.SendPayload(JsonConvert.SerializeObject(package), hostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
                try
                {
                    for (int i = 0; i < _roomItems.Count; i++)
                    {
                        if (_roomItems[i].Room.RoomGuid.Equals(guid))
                        {
                            for (int j = 0; j < _roomItems[i].AgendaItems.Count; j++)
                            {
                                if (_roomItems[i].AgendaItems[j].Id == updatedAgendaItem.Id)
                                {
                                    _roomItems[i].AgendaItems[j] = updatedAgendaItem;
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
                catch (Exception)
                {

                }
            }
        }

        private async Task ProcessPackage(Package package, string hostName)
        {
            if (package != null)
            {
                switch ((PayloadType)package.PayloadType)
                {
                    case PayloadType.Occupancy:
                        break;
                    case PayloadType.Room:
                        break;
                    case PayloadType.Schedule:
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
                        var userProperty = _appointmentItem.UserProperties.Find("RemoteDbEntityId");
                        if (userProperty != null) userProperty.Value = (int)Convert.ChangeType(package.Payload, typeof(int));
                        _appointmentItem.Save();
                        break;
                    case PayloadType.Discovery:
                        break;
                    case PayloadType.PropertyChanged:
                        await _networkCommunication.SendPayload(JsonConvert.SerializeObject(_discoveryPackage), null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
                        break;
                    default:
                        break;
                }
            }
        }

        private string GetHostName(IList<RoomItem> roomItems, Outlook.AppointmentItem appointmentItem)
        {
            if (appointmentItem == null || roomItems == null) return null;
            string guid = appointmentItem.UserProperties.Find("RoomGuid").Value;
            return roomItems.Where(x => x.Room.RoomGuid.Equals(guid)).Select(x => x.HostName).FirstOrDefault();
        }

        //private string GetGuid(string appointmentItemResources)
        //{
        //    if (appointmentItemResources == null) return null;
        //    string guidRegEx = @"[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}";
        //    var resources = appointmentItemResources.Split(';');
        //    foreach (var resource in resources)
        //    {
        //        if (Regex.IsMatch(resource, guidRegEx)) return resource;
        //    }
        //    return null;
        //}

        private Outlook.MAPIFolder CreateCustomCalendarIfNotExists(string calendarName)
        {
            Outlook.MAPIFolder primaryCalendar = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            foreach (Outlook.MAPIFolder personalCalendar in primaryCalendar.Folders)
            {
                if (personalCalendar.Name == calendarName) return personalCalendar;
            }
            return primaryCalendar.Folders.Add(calendarName, Outlook.OlDefaultFolders.olFolderCalendar);
        }

        private void ClearCalendar(RoomItem roomItem)
        {
            Outlook.Folder outlookFolderDeletedItems = (Outlook.Folder)Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
            for (int i = _roomInfocalendar.Items.Count; i > 0; i--)
            {
                if (((string)(_roomInfocalendar.Items[i] as Outlook.AppointmentItem).UserProperties.Find("RoomGuid").Value).Equals(roomItem.Room.RoomGuid))
                {
                    try
                    {
                        int id = (_roomInfocalendar.Items[i] as Outlook.AppointmentItem).UserProperties.Find("RemoteDbEntityId").Value;
                        (_roomInfocalendar.Items[i] as Outlook.AppointmentItem).Delete();
                        for (int j = outlookFolderDeletedItems.Items.Count; j > 0; j--)
                        {
                            try
                            {
                                if (outlookFolderDeletedItems.Items[j].UserProperties.Find("RemoteDbEntityId").Value == id) outlookFolderDeletedItems.Items[j].Delete();
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            for (int i = 0; i < roomItem.AgendaItems.Count; i++) roomItem.AgendaItems[i].Id = -1;
        }

        private void AddItemsToCalendar(List<AgendaItem> agendaItems, Room room)
        {
            if (agendaItems != null && agendaItems.Count > 0)
            {
                try
                {
                    foreach (var agendaItem in agendaItems)
                    {
                        bool cont = true;
                        foreach (Outlook.AppointmentItem item in _roomInfocalendar.Items)
                        {
                            if (agendaItem.Id == item.UserProperties.Find("RemoteDbEntityId").Value) cont = false;
                            if (agendaItem.Start.DateTime == item.Start && agendaItem.End.DateTime == item.End) cont = false;
                        }
                        if (cont)
                        {
                            Outlook.AppointmentItem appointmentItem = _roomInfocalendar.Items.Add(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                            appointmentItem.Location = !(string.IsNullOrEmpty(room.RoomNumber) || string.IsNullOrWhiteSpace(room.RoomNumber))
                            ? !(string.IsNullOrEmpty(room.RoomName) || string.IsNullOrWhiteSpace(room.RoomName)) ? room.RoomName + " " + room.RoomNumber : room.RoomNumber
                            : !(string.IsNullOrEmpty(room.RoomName) || string.IsNullOrWhiteSpace(room.RoomName)) ? room.RoomName : "";
                            appointmentItem.Start = agendaItem.Start.DateTime;
                            appointmentItem.End = agendaItem.End.DateTime;
                            appointmentItem.Subject = agendaItem.Title;
                            appointmentItem.Body = agendaItem.Description;
                            appointmentItem.AllDayEvent = agendaItem.IsAllDayEvent;
                            var remoteDbEntityId = appointmentItem.UserProperties.Add("RemoteDbEntityId", Outlook.OlUserPropertyType.olInteger);
                            var remoteDbEntityTimeStamp = appointmentItem.UserProperties.Add("RemoteDbEntityTimeStamp", Outlook.OlUserPropertyType.olText);
                            var roomGuid = appointmentItem.UserProperties.Add("RoomGuid", Outlook.OlUserPropertyType.olText);
                            var occupancy = appointmentItem.UserProperties.Add("Occupancy", Outlook.OlUserPropertyType.olInteger);
                            remoteDbEntityId.Value = agendaItem.Id;
                            remoteDbEntityTimeStamp.Value = agendaItem.TimeStamp.ToString();
                            roomGuid.Value = room.RoomGuid;
                            occupancy.Value = agendaItem.Occupancy;
                            appointmentItem.Save();
                        }
                    }
                }
                catch (Exception)
                {

                }
            }
        }
    }
}
