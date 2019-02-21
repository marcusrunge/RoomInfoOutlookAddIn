using System;

namespace ModelLibrary
{
    public class AgendaItem
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public DateTimeOffset Start { get; set; }
        public DateTimeOffset End { get; set; }
        public bool IsAllDayEvent { get; set; }
        public bool IsOverridden { get; set; }
        public string Description { get; set; }
        public int Occupancy { get; set; }
        public long TimeStamp { get; set; }
        public bool IsDeleted { get; set; }
    }
}
