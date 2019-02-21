using System;

namespace ModelLibrary
{
    public class PayloadReceivedEventArgs : EventArgs
    {
        public PayloadReceivedEventArgs(string hostName, string package)
        {
            HostName = hostName;
            Package = package;
        }
        public string HostName { get; }
        public string Package { get; }
    }
}
