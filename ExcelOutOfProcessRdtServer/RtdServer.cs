using System;
using System.Runtime.InteropServices;

namespace ExcelOutOfProcessRdtServer
{
    // call with =RTD("ExcelOutOfProcessRdtServer";;"myTopic")
    [ComVisible(true), Guid(Constants.RtdServerClsid), ClassInterface(ClassInterfaceType.None)]
    public sealed class RtdServer : IRtdServer, IDisposable
    {
        private IRTDUpdateEvent? _callback;
        private int _topicId;
        private int _value;

        public event EventHandler<RtdServerEventArgs>? Information;

        public int Value
        {
            get => _value;
            set
            {
                if (_value == value)
                    return;

                _value = value;
                _callback?.UpdateNotify();
            }
        }

        private void OnInformation(string text) => Information?.Invoke(this, new RtdServerEventArgs(text));

        public int ServerStart(IRTDUpdateEvent callback)
        {
            OnInformation("ServerStart called.");
            _callback = callback;
            return 1;
        }

        public object ConnectData(int topicId, ref Array strings, ref bool newValues)
        {
            OnInformation($"ConnectData called with topic id {topicId}. Value {Value} was sent.");
            _topicId = topicId;
            return Value;
        }

        public Array RefreshData(ref int topicCount)
        {
            object[,] data = new object[2, 1];
            data[0, 0] = _topicId;
            data[1, 0] = Value;
            topicCount = 1;
            OnInformation($"RefreshData called. Value {Value} was sent.");
            return data;
        }

        public void DisconnectData(int topicId)
        {
            OnInformation($"DisconnectData called with topic id {topicId}.");
        }

        public int Heartbeat()
        {
            OnInformation("Heartbeat called.");
            return 1;
        }

        public void ServerTerminate()
        {
            OnInformation("ServerTerminate called.");
        }

        public void Dispose() => _callback?.Disconnect();
    }
}
