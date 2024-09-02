using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;

namespace ExcelOutOfProcessRdtServer.Com;

public sealed partial class ComServer : IDisposable
{
    private ConcurrentBag<int> _cookies = [];
    private readonly ConcurrentDictionary<Type, ComClassFactory> _factories = [];

    public void RegisterClassObject<T>(REGCLS options = REGCLS.REGCLS_MULTIPLEUSE, Func<T>? createInstance = null) where T : new()
    {
        if (!_factories.TryGetValue(typeof(T), out var factory))
        {
            createInstance ??= () => new T();
            var instance = createInstance();
            if (instance == null)
                throw new InvalidOperationException();

            factory = new ComClassFactory(instance);
            _factories[typeof(T)] = factory;
        }

        ThrowOnError(CoRegisterClassObject(typeof(T).GUID, factory, CLSCTX.CLSCTX_LOCAL_SERVER, options, out var cookie));
        _cookies.Add(cookie);
    }

    public void Dispose()
    {
        var cookies = Interlocked.Exchange(ref _cookies, []);
        foreach (var cookie in cookies)
        {
            _ = CoRevokeClassObject(cookie);
        }
    }

    public static void Register<T>(RegistryKey rootKey, string? progId = null, string? exePath = null) => Register(rootKey, typeof(T).GUID, progId, exePath);
    public static void Register(RegistryKey rootKey, Guid clsid, string? progId = null, string? exePath = null)
    {
        ArgumentNullException.ThrowIfNull(rootKey);
        ArgumentNullException.ThrowIfNull(clsid);
        if (!string.IsNullOrEmpty(progId))
        {
            using var progidKey = rootKey.CreateSubKey($"Software\\Classes\\{progId}\\CLSID");
            progidKey.SetValue(null, clsid.ToString("B"));
        }

        using var serverKey = rootKey.CreateSubKey($"Software\\Classes\\CLSID\\{clsid:B}\\LocalServer32");
        exePath ??= Environment.ProcessPath ?? throw new ArgumentNullException(nameof(exePath));
        serverKey.SetValue(null, exePath);
    }

    public static bool IsRegistered<T>(RegistryKey rootKey) => IsRegistered(rootKey, typeof(T).GUID);
    public static bool IsRegistered(RegistryKey rootKey, Guid clsid)
    {
        using var key = rootKey.OpenSubKey($"Software\\Classes\\CLSID\\{clsid:B}", false);
        return key != null;
    }

    public static void Unregister<T>(RegistryKey rootKey, string? progId = null) => Unregister(rootKey, typeof(T).GUID, progId);
    public static void Unregister(RegistryKey rootKey, Guid clsid, string? progId = null)
    {
        ArgumentNullException.ThrowIfNull(rootKey);
        ArgumentNullException.ThrowIfNull(clsid);
        rootKey.DeleteSubKeyTree($"Software\\Classes\\CLSID\\{clsid:B}", false);
        if (!string.IsNullOrEmpty(progId))
        {
            rootKey.DeleteSubKeyTree($"Software\\Classes\\{progId}", false);
        }
    }

    public static void Resume() => ThrowOnError(CoResumeClassObjects());
    public static void Suspend() => ThrowOnError(CoSuspendClassObjects());

    internal static void ThrowOnError(int value) { if (value < 0) Marshal.ThrowExceptionForHR(value); }

    [DllImport("ole32")]
    [PreserveSig]
    private static extern int CoRegisterClassObject(in Guid guid, [MarshalAs(UnmanagedType.IUnknown)] object obj, CLSCTX context, REGCLS flags, out int register);

    [DllImport("ole32")]
    [PreserveSig]
    private static extern int CoResumeClassObjects();

    [DllImport("ole32")]
    [PreserveSig]
    private static extern int CoSuspendClassObjects();

    [DllImport("ole32")]
    [PreserveSig]
    private static extern int CoRevokeClassObject(int register);

    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    private class ComClassFactory(object instance) : IClassFactory
    {
        int IClassFactory.LockServer(bool fLock) => 0;
        int IClassFactory.CreateInstance(object? pUnkOuter, in Guid riid, out object? ppvObject)
        {
            if (pUnkOuter != null)
            {
                ppvObject = null;
                const int CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
                return CLASS_E_NOAGGREGATION;
            }

            ppvObject = instance;
            return 0;
        }
    }

    [ComImport, Guid("00000001-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IClassFactory
    {
        [PreserveSig]
        int CreateInstance([MarshalAs(UnmanagedType.IUnknown)] object? pUnkOuter, in Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object? ppvObject);

        [PreserveSig]
        int LockServer(bool fLock);
    }
}
