using System;
using System.Runtime.InteropServices;

namespace ExcelOutOfProcessRdtServer;

[ComImport, Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRTDUpdateEvent
{
    void UpdateNotify();
    int HeartbeatInterval { get; set; }
    void Disconnect();
}
