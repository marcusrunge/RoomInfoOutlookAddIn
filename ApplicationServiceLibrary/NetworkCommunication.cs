using ModelLibrary;
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using static ModelLibrary.Enums;

namespace ApplicationServiceLibrary
{
    public interface INetworkCommunication
    {
        event EventHandler<PayloadReceivedEventArgs> PayloadReceived;
        Task StartConnectionListener(string port, NetworkProtocol networkProtocol);
        Task SendPayload(string payload, string hostName, string port, NetworkProtocol networkProtocol, bool broadcast = false);
    }

    public class NetworkCommunication : INetworkCommunication
    {
        public event EventHandler<PayloadReceivedEventArgs> PayloadReceived;

        public async Task SendPayload(string payload, string hostName, string port, NetworkProtocol networkProtocol, bool broadcast = false)
        {
            switch (networkProtocol)
            {
                case NetworkProtocol.UserDatagram:
                    await SendUserDatagramPayload(payload, hostName, port, broadcast);
                    break;
                case NetworkProtocol.TransmissionControl:
                    await SendTransmissionControlPayload(payload, hostName, port);
                    break;
                default:
                    break;
            }
        }

        public async Task StartConnectionListener(string port, NetworkProtocol networkProtocol)
        {
            switch (networkProtocol)
            {
                case NetworkProtocol.UserDatagram:
                    await ListenForUserDatagramConnection(port);
                    break;
                case NetworkProtocol.TransmissionControl:
                    await ListenForTransmissionControlConnection(port);
                    break;
                default:
                    break;
            }
        }

        private async Task ListenForTransmissionControlConnection(string port)
        {
            TcpListener tcpListener = new TcpListener(IPAddress.Any, int.Parse(port));
            tcpListener.Start();
            while (true)
            {
                try
                {
                    using (TcpClient tcpClient = await tcpListener.AcceptTcpClientAsync())
                    {
                        using (NetworkStream networkStream = tcpClient.GetStream())
                        {
                            StreamReader streamReader = new StreamReader(networkStream, Encoding.UTF8);
                            string response = await streamReader.ReadLineAsync();
                            OnPayloadReceived(new PayloadReceivedEventArgs(((IPEndPoint)tcpClient.Client.RemoteEndPoint).Address.ToString(), response));
                            streamReader.Close();
                            networkStream.Close();
                        }
                        tcpClient.Close();
                    }
                }
                catch { }
            }
        }

        private async Task ListenForUserDatagramConnection(string port)
        {
            UdpClient udpClient = new UdpClient(int.Parse(port));
            while (true)
            {
                try
                {
                    UdpReceiveResult received = await udpClient.ReceiveAsync();
                    OnPayloadReceived(new PayloadReceivedEventArgs(received.RemoteEndPoint.Port.ToString(), Encoding.ASCII.GetString(received.Buffer)));
                }
                catch { }
            }
        }

        private async Task SendUserDatagramPayload(string payload, string hostName, string port, bool broadcast)
        {
            try
            {
                UdpClient udpClient = new UdpClient();
                byte[] bytes = Encoding.ASCII.GetBytes(payload);
                if (broadcast) hostName = "255.255.255.255";
                await udpClient.SendAsync(bytes, bytes.Length, hostName, int.Parse(port));
                udpClient.Close();
            }
            catch { }
        }

        private async Task SendTransmissionControlPayload(string payload, string hostName, string port)
        {
            try
            {
                using (TcpClient tcpClient = new TcpClient())
                {
                    await tcpClient.ConnectAsync(hostName, int.Parse(port));
                    NetworkStream networkStream = tcpClient.GetStream();
                    byte[] payloadAsBytes = Encoding.UTF8.GetBytes(payload + "\n");
                    await networkStream.WriteAsync(payloadAsBytes, 0, payloadAsBytes.Length);
                    using (StreamReader streamReader = new StreamReader(networkStream, Encoding.UTF8))
                    {
                        string response = await streamReader.ReadLineAsync();
                        OnPayloadReceived(new PayloadReceivedEventArgs(((IPEndPoint)tcpClient.Client.RemoteEndPoint).Port.ToString(), response));
                        streamReader.Close();
                        networkStream.Close();
                        tcpClient.Close();
                    }
                }
            }
            catch { }
        }

        void OnPayloadReceived(PayloadReceivedEventArgs e) => PayloadReceived?.Invoke(null, e);
    }
}
