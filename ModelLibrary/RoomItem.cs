using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

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
