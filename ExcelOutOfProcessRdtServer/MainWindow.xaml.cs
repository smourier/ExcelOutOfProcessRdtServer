using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using ExcelOutOfProcessRdtServer.Com;
using Microsoft.Win32;

namespace ExcelOutOfProcessRdtServer;

public partial class MainWindow : Window
{
    private ComServer? _comServer;
    private RtdServer? _server;

    public MainWindow()
    {
        InitializeComponent();
        AppendLog();
        if (Extensions.IsAdministrator)
        {
            Title += " - Administrator";
            AppendLog("In administrator mode, you can only register and unregister the .exe server.");
        }
        else
        {
            RegisterMenuItem.IsEnabled = false;
            UnRegisterMenuItem.IsEnabled = false;
            if (!ComServer.IsRegistered<RtdServer>(Registry.LocalMachine))
            {
                AppendLog("The RdtServer doesn't seem to be registered. Restart this application as admin and register.");
            }
            else
            {
                _comServer = new ComServer();
                _comServer.RegisterClassObject<RtdServer>(createInstance: () =>
                {
                    _server = new RtdServer();
                    _server.Information += OnServerInformation;
                    return _server;
                });

                AppendLog("The Rdt Server has been registered. You can try to open Excel and type something like =RTD(\"ExcelOutOfProcessRdtServer\";;\"myTopic\") in a cell, then press the increment button to see this value change in Excel");
            }
        }
    }

    private void OnServerInformation(object? sender, RtdServerEventArgs e) => Dispatcher.BeginInvoke(() => { AppendLog(e.Information); });

    protected override void OnClosed(EventArgs e)
    {
        Interlocked.Exchange(ref _server, null)?.Dispose();
        Interlocked.Exchange(ref _comServer, null)?.Dispose();
        base.OnClosed(e);
    }

    private void AppendLog(string? message = null)
    {
        Log.Text += message + Environment.NewLine;
    }

    private bool RestartAsAdmin(bool force)
    {
        if (!force && Extensions.IsAdministrator)
            return false;

        var info = new ProcessStartInfo
        {
            FileName = Environment.ProcessPath,
            UseShellExecute = true,
            Verb = "runas" // Provides Run as Administrator
        };

        if (Process.Start(info) != null)
        {
            Close();
            return true;
        }
        return false;
    }

    private void Register()
    {
        ComServer.Register<RtdServer>(Registry.LocalMachine, Constants.RtdServerProgId, Environment.ProcessPath);
        AppendLog("Registration was successful.");
    }

    private void Unregister()
    {
        ComServer.Unregister<RtdServer>(Registry.LocalMachine, Constants.RtdServerProgId);
        AppendLog("Unregistration was successful.");
    }

    private void IncrementName_Click(object sender, RoutedEventArgs e)
    {
        var server = _server;
        if (server != null)
        {
            server.Value++;
            IncrementName.Content = $"Increment Value to {server.Value + 1}";
        }
    }

    private void OnRegister(object sender, RoutedEventArgs e) => Register();
    private void OnUnregister(object sender, RoutedEventArgs e) => Unregister();
    private void OnRestartAsAdmin(object sender, RoutedEventArgs e) => RestartAsAdmin(true);
    private void OnExitClick(object sender, RoutedEventArgs e) => Close();
    private void OnFileOpened(object sender, RoutedEventArgs e)
    {
        var admin = Extensions.IsAdministrator;
        if (!admin)
        {
            if (RestartAsAdminMenuItem.Icon == null)
            {
                const int SHIELD = 77;
                var image = new Image { Source = Extensions.GetStockIconImageSource(SHIELD) };
                RestartAsAdminMenuItem.Icon = image;
            }
            RestartAsAdminMenuItem.Visibility = Visibility.Visible;
            RestartAsAdminSeparator.Visibility = Visibility.Visible;
        }
        else
        {
            RestartAsAdminMenuItem.Visibility = Visibility.Collapsed;
            RestartAsAdminSeparator.Visibility = Visibility.Collapsed;
        }
    }
}