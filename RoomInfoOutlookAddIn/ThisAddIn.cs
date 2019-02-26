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
        Outlook.Items _calendarItems;
        bool _isProcessingPackage;

        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Outlook.MAPIFolder calendar = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            _calendarItems = calendar.Items;
            _calendarItems.ItemAdd += CalendarItems_ItemAdd;
            _calendarItems.ItemChange += CalendarItems_ItemChange;
            _calendarItems.ItemRemove += CalendarItems_ItemRemove;
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
            await _networkCommunication.SendPayload("", null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
        }

        private async void CalendarItems_ItemRemove()
        {
            await _networkCommunication.SendPayload("", null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
            try
            {
                foreach (var roomItem in _roomItems)
                {
                    for (int i = 0; i < roomItem.AgendaItems.Count; i++)
                    {
                        int id = roomItem.AgendaItems[i].Id;
                        string guid = roomItem.Room.RoomGuid;
                        Outlook.AppointmentItem appointmentItem = null;
                        foreach (Outlook.AppointmentItem item in _calendarItems)
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

        private async void _eventService_SyncButtonPressed(object sender, RoomItem roomItem)
        {
            await _networkCommunication.SendPayload("", null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
            //TODO
        }

        private async void ThisAddIn_AddButtonPressed(object sender, RoomItem roomItem)
        {
            await _networkCommunication.SendPayload("", null, Properties.Settings.Default.UdpPort, NetworkProtocol.UserDatagram, true);
            try
            {
                var roomNumber = roomItem.Room.RoomNumber;
                var roomName = roomItem.Room.RoomName;
                Outlook.AppointmentItem newAppointment = (Outlook.AppointmentItem)this.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                newAppointment.Location = !(string.IsNullOrEmpty(roomNumber) || string.IsNullOrWhiteSpace(roomNumber))
                    ? !(string.IsNullOrEmpty(roomName) || string.IsNullOrWhiteSpace(roomName)) ? roomName + " " + roomNumber : roomNumber
                    : !(string.IsNullOrEmpty(roomName) || string.IsNullOrWhiteSpace(roomName)) ? roomName : "";
                newAppointment.Resources = roomItem.Room.RoomGuid;
                var propertyAccessor = newAppointment.PropertyAccessor;
                var userProperty = newAppointment.UserProperties.Add("RemoteDbEntityId", Outlook.OlUserPropertyType.olInteger);
                userProperty.Value = 0;
                newAppointment.Display();
            }
            catch (Exception)
            {

            }
        }

        private async void CalendarItems_ItemChange(object Item)
        {
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
                default:
                    break;
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
    }
}
