namespace ModelLibrary
{
    public class Room
    {
        public string RoomGuid { get; set; }
        public string RoomName { get; set; }
        public string RoomNumber { get; set; }
        public int Occupancy { get; set; }
        public bool IsIoT { get; set; }
    }
}
