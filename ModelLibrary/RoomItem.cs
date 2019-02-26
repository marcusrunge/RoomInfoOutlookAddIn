using System.Collections.Generic;

namespace ModelLibrary
{
    public class RoomItem
    {
        public string HostName { get; set; }
        public List<AgendaItem> AgendaItems { get; set; }
        public StandardWeek StandardWeek { get; set; }
        public Room Room { get; set; }
    }
}
