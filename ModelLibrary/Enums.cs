namespace ModelLibrary
{
    public class Enums
    {
        public enum OccupancyVisualState { FreeVisualState, PresentVisualState, AbsentVisualState, BusyVisualState, OccupiedVisualState, LockedVisualState, UndefinedVisualState }
        public enum PayloadType { Occupancy, Room, Schedule, StandardWeek, RequestOccupancy, RequestSchedule, RequestStandardWeek, IotDim, AgendaItem, AgendaItemId, Discovery, PropertyChanged }
        public enum NetworkProtocol { UserDatagram, TransmissionControl }
    }
}
