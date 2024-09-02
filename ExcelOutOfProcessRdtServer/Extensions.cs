using System;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ExcelOutOfProcessRdtServer;

public static partial class Extensions
{
    public static bool IsAdministrator
    {
        get
        {
            var identity = WindowsIdentity.GetCurrent();
            return identity != null && new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
        }
    }

    public static ImageSource? GetStockIconImageSource(int id)
    {
        var hicon = GetStockIconHandle(id);
        if (hicon == IntPtr.Zero)
            return null;

        try
        {
            return Imaging.CreateBitmapSourceFromHIcon(hicon, new Int32Rect(0, 0, 16, 16), BitmapSizeOptions.FromEmptyOptions());
        }
        finally
        {
            _ = DestroyIcon(hicon);
        }
    }

    public static nint GetStockIconHandle(int id)
    {
        var info = new SHSTOCKICONINFO { cbSize = Marshal.SizeOf<SHSTOCKICONINFO>() };
        const int SHGSI_ICON = 0x100;
        const int SHGSI_SMALLICON = 1;
        _ = SHGetStockIconInfo(id, SHGSI_ICON | SHGSI_SMALLICON, ref info);
        return info.hIcon;
    }

    [DllImport("user32")]
    private static extern int DestroyIcon(nint handle);

    [DllImport("shell32")]
    private static extern int SHGetStockIconInfo(int siid, int uFlags, ref SHSTOCKICONINFO psii);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct SHSTOCKICONINFO
    {
        public int cbSize;
        public IntPtr hIcon;
        public int iSysIconIndex;
        public int iIcon;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szPath;
    }
}
