using ModelLibrary;
using System;

namespace ApplicationServiceLibrary
{
    public interface IEventService
    {
        event EventHandler<RoomItem> AddButtonPressed;
        void OnAddButtonPressed(RoomItem roomItem);

        event EventHandler<RoomItem> SyncButtonPressed;
        void OnSyncButtonPressed(RoomItem roomItem);
    }

    public class EventService : IEventService
    {
        public event EventHandler<RoomItem> AddButtonPressed;
        public void OnAddButtonPressed(RoomItem roomItem) => AddButtonPressed?.Invoke(null, roomItem);

        public event EventHandler<RoomItem> SyncButtonPressed;
        public void OnSyncButtonPressed(RoomItem roomItem) => SyncButtonPressed?.Invoke(null, roomItem);
    }
}
