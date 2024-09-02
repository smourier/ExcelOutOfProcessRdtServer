using System;

namespace ExcelOutOfProcessRdtServer;

public class RtdServerEventArgs(string information) : EventArgs
{
    public string Information { get; } = information;
}
