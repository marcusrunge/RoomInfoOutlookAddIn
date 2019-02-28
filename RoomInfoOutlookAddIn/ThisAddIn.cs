using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using System.Globalization;
using Unity;
using Microsoft.Office.Core;
using ApplicationServiceLibrary;
using ModelLibrary;
using static ModelLibrary.Enums;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace RoomInfoOutlookAddIn
{
    public partial class ThisAddIn
    {
        IUnityContainer _unityContainer;
        IEventService _eventService;
        INetworkCommunication _networkCommunication;
        IList<RoomItem> _roomItems;
        Outlook.AppointmentItem _appointmentItem;
        Outlook.MAPIFolder _calendar;
        bool _isProcessingPackage;
        Package _discoveryPackage;
        bool _isSyncInProgress;

        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _calendar = CreateCustomCalendarIfNotExists("RoomInfoCalendar");
            _calendar.Items.ItemAdd += CalendarItems_ItemAdd;
            _calendar.Items.ItemChange += CalendarItems_ItemChange;
            _calendar.Items.ItemRemove += CalendarItems_ItemRemove;
            _eventService = _unityContainer.Resolve<IEventService>();
            _networkCommunication = _unityContainer.Resolve<INetworkCommunication>();
            _roomItems = _unityContainer.Resolve<IList<RoomItem>>();
            _eventService.AddButtonPressed += ThisAddIn_AddButtonPressed;
            _eventService.SyncButtonPressed += _eventService_SyncButtonPressed;
            await _networkCommunication.StartConnectionListener(Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
            _networkCommunication.PayloadReceived += async (s, ea) =>
            {
                if (ea.Package != null) await ProcessPackage(JsonConvert.DeserializeObject<Package>(ea.Package), ea.HostName);
            };
            await _networkCommunication.SendPayload(JsonConvert.SerializeObject(_discoveryPackage), null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
        }

        private async void CalendarItems_ItemRemove()
        {
            await _networkCommunication.SendPayload(JsonConvert.SerializeObject(_discoveryPackage), null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
            try
            {
                foreach (var roomItem in _roomItems)
                {
                    for (int i = 0; i < roomItem.AgendaItems.Count; i++)
                    {
                        int id = roomItem.AgendaItems[i].Id;
                        string guid = roomItem.Room.RoomGuid;
                        Outlook.AppointmentItem appointmentItem = null;
                        foreach (Outlook.AppointmentItem item in _calendar.Items)
                        {
                            if (item.Resources == null) break;
                            string itemGuid = GetGuid(item.Resources);
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
                    var schedulePackage = new Package() { PayloadType = (int)PayloadType.RequestSchedule };
                    await _networkCommunication.SendPayload(JsonConvert.SerializeObject(schedulePackage), roomItem.HostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
                }
            }
            catch (Exception)
            {

            }
        }

        private void _eventService_SyncButtonPressed(object sender, RoomItem roomItem)
        {
            _isSyncInProgress = true;
            ClearCalendar(roomItem.Room.RoomGuid);
            AddItemsToCalendar(roomItem.AgendaItems, roomItem.Room);
            _isSyncInProgress = false;
        }

        private async void ThisAddIn_AddButtonPressed(object sender, RoomItem roomItem)
        {
            await _networkCommunication.SendPayload("", null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
            try
            {
                var roomNumber = roomItem.Room.RoomNumber;
                var roomName = roomItem.Room.RoomName;
                Outlook.AppointmentItem newAppointment = _calendar.Items.Add(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                newAppointment.Location = !(string.IsNullOrEmpty(roomNumber) || string.IsNullOrWhiteSpace(roomNumber))
                    ? !(string.IsNullOrEmpty(roomName) || string.IsNullOrWhiteSpace(roomName)) ? roomName + " " + roomNumber : roomNumber
                    : !(string.IsNullOrEmpty(roomName) || string.IsNullOrWhiteSpace(roomName)) ? roomName : "";
                newAppointment.Resources = roomItem.Room.RoomGuid;
                var remoteDbEntityId = newAppointment.UserProperties.Add("RemoteDbEntityId", Outlook.OlUserPropertyType.olInteger);
                var remoteDbEntityTimeStamp = newAppointment.UserProperties.Add("RemoteDbEntityTimeStamp", Outlook.OlUserPropertyType.olDuration);
                remoteDbEntityId.Value = 0;
                remoteDbEntityTimeStamp.Value = DateTimeOffset.Now.ToUnixTimeMilliseconds();
                newAppointment.Display();
            }
            catch (Exception)
            {

            }
        }

        private async void CalendarItems_ItemChange(object Item)
        {
            if (_isSyncInProgress) return;
            if (_isProcessingPackage)
            {
                _isProcessingPackage = false;
                return;
            }
            _appointmentItem = Item as Outlook.AppointmentItem;
            await TransmitAgendaItem(_appointmentItem);
            var schedulePackage = new Package() { PayloadType = (int)PayloadType.RequestSchedule };
            await _networkCommunication.SendPayload(JsonConvert.SerializeObject(schedulePackage), GetHostName(_roomItems, _appointmentItem), Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
        }

        private async void CalendarItems_ItemAdd(object Item)
        {
            if (_isSyncInProgress) return;
            _appointmentItem = Item as Outlook.AppointmentItem;
            await TransmitAgendaItem(_appointmentItem);
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
            _isProcessingPackage = false;
            _unityContainer = new UnityContainer();
            _unityContainer.RegisterSingleton<IMainRibbon, MainRibbon>();
            _unityContainer.RegisterSingleton<INetworkCommunication, NetworkCommunication>();
            _unityContainer.RegisterSingleton<IEventService, EventService>();
            _unityContainer.RegisterSingleton<IList<RoomItem>, List<RoomItem>>();
            Outlook.Application outlookApplication = GetHostItem<Outlook.Application>(typeof(Outlook.Application), "Application");
            int languageID = outlookApplication.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(languageID);
            return _unityContainer.Resolve<IMainRibbon>();
        }

        private async Task TransmitAgendaItem(Outlook.AppointmentItem appointmentItem)
        {
            string hostName = GetHostName(_roomItems, appointmentItem);
            var agendaItem = new AgendaItem()
            {
                Id = appointmentItem.UserProperties.Find("RemoteDbEntityId").Value,
                Description = appointmentItem.Body,
                End = appointmentItem.End,
                IsAllDayEvent = appointmentItem.AllDayEvent,
                Start = appointmentItem.Start,
                Title = appointmentItem.Subject,
                Occupancy = 3
            };
            string guid = GetGuid(appointmentItem.Resources);
            var agendaItems = _roomItems.Where(x => x.Room.RoomGuid.Equals(guid)).Select(x => x.AgendaItems).FirstOrDefault();
            if (agendaItems != null)
            {
                var actualAgendaItem = agendaItems.Where(x => x.Id == agendaItem.Id).Select(x => x).FirstOrDefault();
                if (actualAgendaItem != null && actualAgendaItem.Equals(agendaItem)) return;
                var package = new Package() { PayloadType = (int)PayloadType.AgendaItem, Payload = agendaItem };
                await _networkCommunication.SendPayload(JsonConvert.SerializeObject(package), hostName, Properties.Settings.Default.TcpPort, NetworkProtocol.TransmissionControl);
            }
        }

        private Task ProcessPackage(Package package, string hostName)
        {
            _isProcessingPackage = true;
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
                        break;
                    default:
                        break;
                }
            }
            return Task.CompletedTask;
        }

        private string GetHostName(IList<RoomItem> roomItems, Outlook.AppointmentItem appointmentItem)
        {
            string guid = GetGuid(appointmentItem.Resources);
            return roomItems.Where(x => x.Room.RoomGuid.Equals(guid)).Select(x => x.HostName).FirstOrDefault();
        }

        private string GetGuid(string appointmentItemResources)
        {
            string guidRegEx = @"[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}";
            var resources = appointmentItemResources.Split(';');
            foreach (var resource in resources)
            {
                if (Regex.IsMatch(resource, guidRegEx)) return resource;
            }
            return null;
        }

        private Outlook.MAPIFolder CreateCustomCalendarIfNotExists(string calendarName)
        {
            Outlook.MAPIFolder primaryCalendar = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            foreach (Outlook.MAPIFolder personalCalendar in primaryCalendar.Folders)
            {
                if (personalCalendar.Name == calendarName) return personalCalendar;
            }
            return primaryCalendar.Folders.Add(calendarName, Outlook.OlDefaultFolders.olFolderCalendar);
        }

        private void ClearCalendar(string roomGuid)
        {
            int c = _calendar.Items.Count;
            for (int i = _calendar.Items.Count; i > 0; i--)
            {
                if (GetGuid((_calendar.Items[i] as Outlook.AppointmentItem).Resources).Equals(roomGuid))
                    (_calendar.Items[i] as Outlook.AppointmentItem).Delete();
            }
        }

        private void AddItemsToCalendar(List<AgendaItem> agendaItems, Room room)
        {
            if (agendaItems != null && agendaItems.Count > 0)
            {
                foreach (var agendaItem in agendaItems)
                {
                    Outlook.AppointmentItem appointmentItem = _calendar.Items.Add(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                    appointmentItem.Location = !(string.IsNullOrEmpty(room.RoomNumber) || string.IsNullOrWhiteSpace(room.RoomNumber))
                    ? !(string.IsNullOrEmpty(room.RoomName) || string.IsNullOrWhiteSpace(room.RoomName)) ? room.RoomName + " " + room.RoomNumber : room.RoomNumber
                    : !(string.IsNullOrEmpty(room.RoomName) || string.IsNullOrWhiteSpace(room.RoomName)) ? room.RoomName : "";
                    appointmentItem.Start = agendaItem.Start.DateTime;
                    appointmentItem.End = agendaItem.End.DateTime;
                    appointmentItem.Subject = agendaItem.Title;
                    appointmentItem.Body = agendaItem.Description;
                    appointmentItem.AllDayEvent = agendaItem.IsAllDayEvent;
                    appointmentItem.Resources = room.RoomGuid;
                    var remoteDbEntityId = appointmentItem.UserProperties.Add("RemoteDbEntityId", Outlook.OlUserPropertyType.olInteger);
                    //var remoteDbEntityTimeStamp = appointmentItem.UserProperties.Add("RemoteDbEntityTimeStamp", Outlook.OlUserPropertyType.olNumber);
                    remoteDbEntityId.Value = agendaItem.Id;
                    //remoteDbEntityTimeStamp.Value = DateTimeOffset.Now.ToUnixTimeMilliseconds();
                    appointmentItem.Save();
                }
            }
        }
    }
}
