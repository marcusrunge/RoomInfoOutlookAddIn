using ModelLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApplicationServiceLibrary
{
    public interface IEventService
    {
        event EventHandler<RoomItem> AddButtonPressed;
        void OnAddButtonPressed(RoomItem roomItem);
    }

    public class EventService : IEventService
    {
        public event EventHandler<RoomItem> AddButtonPressed;
        public void OnAddButtonPressed(RoomItem roomItem) => AddButtonPressed?.Invoke(null, roomItem);
    }
}
