# Pro-Integrate
First version of a “Integrate Apps for Dummies” that specifically is designed to help pull all 3rd party apps, along with Microsoft Store apps into one easy to use panel where the user will be able to see which apps can be integrated with what, to make all apps work together instead of having to rely on multi-Tabs. Truly a One Network.
<# ======================================================================
  bootstrap.ps1  —  One-drop repo generator for Company Office Integrator
  Usage (from repo root):
    powershell -ExecutionPolicy Bypass -File .\bootstrap.ps1
====================================================================== #>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = (Get-Location).Path

function New-TextFile {
  param(
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)][string]$Content
  )
  $dir = Split-Path -Parent $Path
  if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  $Content | Out-File -FilePath $Path -Encoding UTF8 -Force
}

Write-Host "==> Creating solution skeleton..." -ForegroundColor Cyan

# Ensure .NET SDK is available
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
  throw "dotnet SDK not found. Install .NET 8 SDK first."
}

# Initialize git if not already
if (-not (Test-Path "$root\.git")) {
  git init | Out-Null
}

# Create solution
& dotnet new sln -n DesktopIntegrator | Out-Null

# Create projects
& dotnet new wpf -n Integrator.UI.Wpf -o src/Integrator.UI.Wpf --framework net8.0-windows | Out-Null
& dotnet new classlib -n Integrator.Engine -o src/Integrator.Engine --framework net8.0-windows | Out-Null
& dotnet new classlib -n Integrator.Shared -o src/Integrator.Shared --framework net8.0 | Out-Null
& dotnet new console -n Integrator.Cli -o src/Integrator.Cli --framework net8.0 | Out-Null
& dotnet new console -n Integrator.Helper.Admin -o src/Integrator.Helper.Admin --framework net8.0-windows | Out-Null

# Add to solution
& dotnet sln DesktopIntegrator.sln add src/Integrator.UI.Wpf/Integrator.UI.Wpf.csproj | Out-Null
& dotnet sln DesktopIntegrator.sln add src/Integrator.Engine/Integrator.Engine.csproj | Out-Null
& dotnet sln DesktopIntegrator.sln add src/Integrator.Shared/Integrator.Shared.csproj | Out-Null
& dotnet sln DesktopIntegrator.sln add src/Integrator.Cli/Integrator.Cli.csproj | Out-Null
& dotnet sln DesktopIntegrator.sln add src/Integrator.Helper.Admin/Integrator.Helper.Admin.csproj | Out-Null

# Overwrite csproj files with tuned settings
New-TextFile -Path src/Integrator.UI.Wpf/Integrator.UI.Wpf.csproj -Content @'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net8.0-windows10.0.19041.0</TargetFramework>
    <UseWPF>true</UseWPF>
    <Platforms>x64</Platforms>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <AssemblyName>Integrator.UI.Wpf</AssemblyName>
    <RootNamespace>Integrator.UI.Wpf</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\Integrator.Engine\Integrator.Engine.csproj" />
    <ProjectReference Include="..\Integrator.Shared\Integrator.Shared.csproj" />
  </ItemGroup>
  <ItemGroup>
    <None Update="..\..\payload\AnalyticsManifest.xml">
      <Link>payload\AnalyticsManifest.xml</Link>
      <CopyToOutputDirectory>Always</CopyToOutputDirectory>
    </None>
  </ItemGroup>
</Project>
'@

New-TextFile -Path src/Integrator.Engine/Integrator.Engine.csproj -Content @'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <Platforms>x64</Platforms>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.Data.Sqlite" Version="8.*" />
    <PackageReference Include="Serilog" Version="3.*" />
    <PackageReference Include="Serilog.Sinks.File" Version="5.*" />
    <PackageReference Include="Microsoft.Identity.Client" Version="4.*" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\Integrator.Shared\Integrator.Shared.csproj" />
  </ItemGroup>
</Project>
'@

New-TextFile -Path src/Integrator.Shared/Integrator.Shared.csproj -Content @'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <Platforms>x64</Platforms>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>
</Project>
'@

New-TextFile -Path src/Integrator.Cli/Integrator.Cli.csproj -Content @'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <Platforms>x64</Platforms>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="System.CommandLine" Version="2.0.0-beta4.22272.1" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\Integrator.Engine\Integrator.Engine.csproj" />
    <ProjectReference Include="..\Integrator.Shared\Integrator.Shared.csproj" />
  </ItemGroup>
</Project>
'@

New-TextFile -Path src/Integrator.Helper.Admin/Integrator.Helper.Admin.csproj -Content @'
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0-windows</TargetFramework>
    <Platforms>x64</Platforms>
    <Nullable>enable</Nullable>
    <ApplicationManifest>App.manifest</ApplicationManifest>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\Integrator.Shared\Integrator.Shared.csproj" />
  </ItemGroup>
</Project>
'@

# --- UI files ---
New-TextFile -Path src/Integrator.UI.Wpf/App.xaml -Content @'
<Application x:Class="Integrator.UI.Wpf.App"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             StartupUri="MainWindow.xaml">
  <Application.Resources/>
</Application>
'@

New-TextFile -Path src/Integrator.UI.Wpf/App.xaml.cs -Content @'
using System.Windows;

namespace Integrator.UI.Wpf;
public partial class App : Application { }
'@

New-TextFile -Path src/Integrator.UI.Wpf/MainWindow.xaml -Content @'
<Window x:Class="Integrator.UI.Wpf.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Company Office Integrator" Height="520" Width="840">
  <DockPanel Margin="16">
    <StackPanel Orientation="Horizontal" DockPanel.Dock="Top" Margin="0,0,0,12">
      <CheckBox Content="Integrate All" IsChecked="{Binding IntegrateAll}" Margin="0,0,12,0"/>
      <Button Content="Run" Width="110" Command="{Binding RunCommand}"/>
      <Button Content="Rollback" Width="110" Margin="8,0,0,0" Command="{Binding RollbackCommand}"/>
      <CheckBox Content="Machine Scope (UAC)" Margin="24,0,0,0" IsChecked="{Binding MachineScope}"/>
    </StackPanel>
    <Border BorderBrush="Gray" BorderThickness="1" Padding="8">
      <StackPanel>
        <ItemsControl ItemsSource="{Binding Hosts}">
          <ItemsControl.ItemTemplate>
            <DataTemplate>
              <StackPanel Orientation="Horizontal" Margin="0,6">
                <CheckBox IsChecked="{Binding IsSelected}" Width="20"/>
                <TextBlock Text="{Binding Name}" Width="140"/>
                <TextBlock Text="{Binding Status}" />
              </StackPanel>
            </DataTemplate>
          </ItemsControl.ItemTemplate>
        </ItemsControl>
        <TextBlock Text="Log:" Margin="0,12,0,0" FontWeight="Bold"/>
        <TextBox Text="{Binding Log}" VerticalScrollBarVisibility="Auto" AcceptsReturn="True" Height="220" />
      </StackPanel>
    </Border>
  </DockPanel>
</Window>
'@

New-TextFile -Path src/Integrator.UI.Wpf/MainWindow.xaml.cs -Content @'
using System.Windows;

namespace Integrator.UI.Wpf;
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new IntegrateViewModel();
    }
}
'@

New-TextFile -Path src/Integrator.UI.Wpf/RelayCommand.cs -Content @'
using System;
using System.Windows.Input;

namespace Integrator.UI.Wpf;
public sealed class RelayCommand : ICommand
{
    private readonly Action<object?> _exec;
    private readonly Func<object?, bool>? _can;
    public RelayCommand(Action<object?> exec, Func<object?, bool>? can = null) { _exec = exec; _can = can; }
    public bool CanExecute(object? p) => _can?.Invoke(p) ?? true;
    public void Execute(object? p) => _exec(p);
    public event EventHandler? CanExecuteChanged;
    public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}
'@

New-TextFile -Path src/Integrator.UI.Wpf/IntegrateViewModel.cs -Content @'
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows.Input;
using Integrator.Engine;

namespace Integrator.UI.Wpf;

public sealed class IntegrateViewModel : INotifyPropertyChanged
{
    public ObservableCollection<HostItem> Hosts { get; } = new()
    {
        new HostItem("Word"),
        new HostItem("Excel"),
        new HostItem("PowerPoint"),
        new HostItem("Outlook")
    };

    private bool _integrateAll;
    public bool IntegrateAll
    {
        get => _integrateAll;
        set { _integrateAll = value; foreach (var h in Hosts) h.IsSelected = value; OnPropertyChanged(); }
    }

    private bool _machineScope;
    public bool MachineScope { get => _machineScope; set { _machineScope = value; OnPropertyChanged(); } }

    private string _log = "";
    public string Log { get => _log; set { _log = value; OnPropertyChanged(); } }

    public ICommand RunCommand { get; }
    public ICommand RollbackCommand { get; }

    // Sample add-in info
    private const string VstoAddinId = "Company.Reports";
    private const string VstoFriendly = "Company Reports";
    private const string VstoDesc = "Reporting tools";
    private static readonly string VstoManifestPath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                     "Company","Addins","Reports","Reports.vsto");

    private static readonly string CatalogId = "{b0532b67-2f2e-4c2f-b87c-1f1c2b9f85a1}";
    private static readonly string CatalogPath = Paths.CatalogPath;
    private const string WebAddinId = "Company.Analytics";
    private readonly string _webManifestPathOnDisk =
        Path.Combine(AppContext.BaseDirectory, "payload", "AnalyticsManifest.xml");

    public IntegrateViewModel()
    {
        RunCommand = new RelayCommand(_ => Run());
        RollbackCommand = new RelayCommand(_ => Rollback());
        Directory.CreateDirectory(Path.GetDirectoryName(VstoManifestPath)!); // ensure folder for demo
        // Drop a placeholder Reports.vsto if none exists (so registry path is valid for smoke tests)
        var vsto = VstoManifestPath;
        if (!File.Exists(vsto))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(vsto)!);
            File.WriteAllText(vsto, "<deployment/>");
        }
    }

    private void Run()
    {
        var selected = Hosts.Where(h => h.IsSelected).ToList();
        if (selected.Count == 0) { Append("Nothing selected."); return; }

        var txn = new MutationTransaction(Paths.RollbackDir);
        try
        {
            Append($"Starting integration. Scope={(MachineScope ? "machine" : "user")}.");

            foreach (var host in selected)
            {
                host.Status = "Integrating...";
                AddinRegistrarVstoCom.RegisterVsto(
                    txn, MachineScope, host.Name, VstoAddinId, VstoFriendly, VstoDesc, VstoManifestPath);
                Append($"VSTO registered for {host.Name}");
            }

            AddinRegistrarWeb.EnsureTrustedCatalog(txn, CatalogPath, CatalogId);
            var manifest = File.ReadAllText(_webManifestPathOnDisk, Encoding.UTF8);
            AddinRegistrarWeb.AddManifest(txn, CatalogPath, WebAddinId, manifest);
            Append($"Web add-in manifest '{WebAddinId}.xml' written to {CatalogPath}");

            txn.Commit();
            foreach (var host in selected) host.Status = "Integrated ✅";
            Append("Integration complete.");
        }
        catch (Exception ex)
        {
            Append("ERROR: " + ex.Message);
            try { txn.Rollback(); Append("Rolled back changes."); } catch { }
            foreach (var host in selected) host.Status = "Failed ❌";
        }
    }

    private void Rollback()
    {
        var selected = Hosts.Where(h => h.IsSelected).ToList();
        if (selected.Count == 0) { Append("Nothing selected."); return; }

        var txn = new MutationTransaction(Paths.RollbackDir);
        try
        {
            Append($"Rolling back selections. Scope={(MachineScope ? "machine" : "user")}.");

            foreach (var host in selected)
            {
                AddinRegistrarVstoCom.Unregister(txn, MachineScope, host.Name, VstoAddinId);
                host.Status = "Unregistered VSTO";
            }

            AddinRegistrarWeb.RemoveManifest(txn, CatalogPath, WebAddinId);
            Append($"Removed '{WebAddinId}.xml' from catalog.");

            txn.Commit();
            foreach (var host in selected) host.Status = "Rolled back ⬅";
            Append("Rollback complete.");
        }
        catch (Exception ex)
        {
            Append("ERROR (rollback): " + ex.Message);
            try { txn.Rollback(); Append("Rollback reverted."); } catch { }
        }
    }

    private void Append(string line) => Log += $"[{DateTime.Now:T}] {line}{Environment.NewLine}";

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? p = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
}

public sealed class HostItem : INotifyPropertyChanged
{
    public string Name { get; }
    private bool _isSelected;
    private string _status = "Idle";

    public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(); } }
    public string Status { get => _status; set { _status = value; OnPropertyChanged(); } }

    public HostItem(string name) { Name = name; }
    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? p = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
}
'@

# --- Engine files ---
New-TextFile -Path src/Integrator.Engine/MutationTransaction.cs -Content @'
using Microsoft.Win32;
using System.Text.Json;

namespace Integrator.Engine;

public sealed class MutationTransaction : IDisposable
{
    private readonly List<IUndoable> _ops = new();
    private readonly string _logDir;
    private readonly string _logFile;

    public MutationTransaction(string logDir)
    {
        _logDir = logDir;
        Directory.CreateDirectory(_logDir);
        _logFile = Path.Combine(_logDir, $"txn_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.json");
    }

    // Registry set
    public void SetRegValue(RegistryHive hive, string subKey, string name, object? value, RegistryValueKind kind)
    {
        var root = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
        using var key = root.CreateSubKey(subKey, true) ?? throw new IOException($"Cannot open {hive}\\{subKey}");
        object? before = null;
        var hadValue = key.GetValueNames().Contains(name);
        if (hadValue) before = key.GetValue(name);

        _ops.Add(new RegSet(hive, subKey, name, hadValue, before));
        key.SetValue(name, value ?? string.Empty, kind);
    }

    // Registry subtree delete
    public void DeleteRegKeyTree(RegistryHive hive, string subKey)
    {
        var root = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
        using var key = root.OpenSubKey(subKey);
        if (key == null) return;

        var snapshot = SnapshotKey(hive, subKey);
        _ops.Add(new RegDeleteTree(hive, subKey, snapshot));
        root.DeleteSubKeyTree(subKey, false);
    }

    // File write helpers
    public void CopyFileWithBackup(string src, string dst, bool overwrite = true)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(dst)!);
        byte[]? before = File.Exists(dst) ? File.ReadAllBytes(dst) : null;
        _ops.Add(new FileWrite(dst, before));
        File.Copy(src, dst, overwrite);
    }

    public void WriteText(string dst, string content)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(dst)!);
        byte[]? before = File.Exists(dst) ? File.ReadAllBytes(dst) : null;
        _ops.Add(new FileWrite(dst, before));
        File.WriteAllText(dst, content);
    }

    public void Commit()
    {
        var json = JsonSerializer.Serialize(_ops, new JsonSerializerOptions { WriteIndented = true, IncludeFields = true });
        File.WriteAllText(_logFile, json);
    }

    public void Rollback()
    {
        for (int i = _ops.Count - 1; i >= 0; i--) _ops[i].Undo();
        File.AppendAllText(_logFile, $"{Environment.NewLine}/* ROLLBACK @{DateTime.UtcNow:o} */");
    }

    public void Dispose() { }

    private static RegistrySnapshot SnapshotKey(RegistryHive hive, string subKey)
    {
        var root = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
        using var k = root.OpenSubKey(subKey);
        var snapshot = new RegistrySnapshot { Hive = hive, SubKey = subKey };
        if (k == null) return snapshot;

        foreach (var n in k.GetValueNames())
            snapshot.Values[n] = k.GetValue(n)?.ToString();

        foreach (var child in k.GetSubKeyNames())
            snapshot.Children.Add(SnapshotKey(hive, $"{subKey}\\{child}"));

        return snapshot;
    }

    private interface IUndoable { void Undo(); }

    private record RegSet(RegistryHive Hive, string SubKey, string Name, bool HadValue, object? Before) : IUndoable
    {
        public void Undo()
        {
            var root = RegistryKey.OpenBaseKey(Hive, RegistryView.Registry64);
            using var key = root.CreateSubKey(SubKey, true);
            if (key == null) return;
            if (HadValue) key.SetValue(Name, Before ?? "");
            else key.DeleteValue(Name, false);
        }
    }

    private record RegDeleteTree(RegistryHive Hive, string SubKey, RegistrySnapshot Snapshot) : IUndoable
    {
        public void Undo()
        {
            var root = RegistryKey.OpenBaseKey(Hive, RegistryView.Registry64);
            Restore(root, Snapshot);
        }
        private static void Restore(RegistryKey root, RegistrySnapshot s)
        {
            using var key = root.CreateSubKey(s.SubKey, true);
            foreach (var kv in s.Values) key?.SetValue(kv.Key, kv.Value);
            foreach (var c in s.Children) Restore(root, c);
        }
    }

    private record FileWrite(string PathToFile, byte[]? Before) : IUndoable
    {
        public void Undo()
        {
            if (Before == null) { if (File.Exists(PathToFile)) File.Delete(PathToFile); }
            else { Directory.CreateDirectory(System.IO.Path.GetDirectoryName(PathToFile)!); File.WriteAllBytes(PathToFile, Before); }
        }
    }

    private class RegistrySnapshot
    {
        public RegistryHive Hive { get; set; }
        public string SubKey { get; set; } = "";
        public Dictionary<string, string?> Values { get; } = new();
        public List<RegistrySnapshot> Children { get; } = new();
    }
}
'@

New-TextFile -Path src/Integrator.Engine/AddinRegistrar.VstoCom.cs -Content @'
using Microsoft.Win32;

namespace Integrator.Engine;

public static class AddinRegistrarVstoCom
{
    public static void RegisterVsto(MutationTransaction txn, bool machineScope, string app, string addinId,
        string friendlyName, string description, string vstoManifestPath)
    {
        var hive = machineScope ? RegistryHive.LocalMachine : RegistryHive.CurrentUser;
        var key = $@"Software\Microsoft\Office\{app}\Addins\{addinId}";
        txn.SetRegValue(hive, key, "FriendlyName", friendlyName, RegistryValueKind.String);
        txn.SetRegValue(hive, key, "Description", description, RegistryValueKind.String);
        txn.SetRegValue(hive, key, "LoadBehavior", 3, RegistryValueKind.DWord);
        txn.SetRegValue(hive, key, "Manifest", vstoManifestPath, RegistryValueKind.String);
    }

    public static void RegisterCom(MutationTransaction txn, bool machineScope, string app, string addinId,
        string friendlyName, string description, string progId)
    {
        var hive = machineScope ? RegistryHive.LocalMachine : RegistryHive.CurrentUser;
        var key = $@"Software\Microsoft\Office\{app}\Addins\{addinId}";
        txn.SetRegValue(hive, key, "FriendlyName", friendlyName, RegistryValueKind.String);
        txn.SetRegValue(hive, key, "Description", description, RegistryValueKind.String);
        txn.SetRegValue(hive, key, "LoadBehavior", 3, RegistryValueKind.DWord);
        txn.SetRegValue(hive, key, "CommandLineSafe", 0, RegistryValueKind.DWord);
        // COM ProgID/CLSID registration should exist separately.
    }

    public static void Unregister(MutationTransaction txn, bool machineScope, string app, string addinId)
    {
        var hive = machineScope ? RegistryHive.LocalMachine : RegistryHive.CurrentUser;
        var key = $@"Software\Microsoft\Office\{app}\Addins\{addinId}";
        txn.DeleteRegKeyTree(hive, key);
    }
}
'@

New-TextFile -Path src/Integrator.Engine/AddinRegistrar.Web.cs -Content @'
using Microsoft.Win32;

namespace Integrator.Engine;

public static class AddinRegistrarWeb
{
    private const string WefTrusted = @"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs";

    public static string EnsureTrustedCatalog(MutationTransaction txn, string catalogPath, string catalogId)
    {
        Directory.CreateDirectory(catalogPath);
        var key = $@"{WefTrusted}\{catalogId}";
        var url = new Uri(catalogPath).AbsoluteUri.TrimEnd('/') + "/"; // file:///C:/ProgramData/.../
        txn.SetRegValue(RegistryHive.CurrentUser, key, "Url", url, RegistryValueKind.String);
        txn.SetRegValue(RegistryHive.CurrentUser, key, "Flags", 1, RegistryValueKind.DWord);
        return url;
    }

    public static void AddManifest(MutationTransaction txn, string catalogPath, string id, string manifestXmlContent)
    {
        var dst = Path.Combine(catalogPath, $"{id}.xml");
        txn.WriteText(dst, manifestXmlContent);
    }

    public static void RemoveManifest(MutationTransaction txn, string catalogPath, string id)
    {
        var dst = Path.Combine(catalogPath, $"{id}.xml");
        if (File.Exists(dst))
        {
            // record previous content and remove
            txn.WriteText(dst, string.Empty);
            File.Delete(dst);
        }
    }
}
'@

New-TextFile -Path src/Integrator.Engine/Paths.cs -Content @'
namespace Integrator.Engine;
public static class Paths
{
    public static string LocalRoot => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Company","Integrator");
    public static string RollbackDir => Path.Combine(LocalRoot, "rollback");
    public static string PerUserAddinsRoot => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Company","Addins");
    public static string PerMachineAddinsRoot => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        "Company","Addins");
    public static string CatalogPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
        "Company","AddinCatalog"); // %ProgramData%
}
'@

# --- Shared stubs ---
New-TextFile -Path src/Integrator.Shared/Contracts.cs -Content @'
namespace Integrator.Shared;
public record ModuleSelection(string ProfileId, string ModuleId, string App, string Scope);
'@

New-TextFile -Path src/Integrator.Shared/Logging.cs -Content @'
using Serilog;

namespace Integrator.Shared;
public static class LogInit
{
    public static void Init(string logDir)
    {
        Directory.CreateDirectory(logDir);
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(Path.Combine(logDir, "integrator-.log"),
                          rollingInterval: RollingInterval.Day,
                          retainedFileCountLimit: 7,
                          shared: true)
            .CreateLogger();
    }
}
'@

New-TextFile -Path src/Integrator.Shared/Storage.cs -Content @'
using Microsoft.Data.Sqlite;

namespace Integrator.Shared;
public static class Storage
{
    public static string DbPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Company","Integrator","state.db");

    public static void Ensure()
    {
        var dir = Path.GetDirectoryName(DbPath)!;
        Directory.CreateDirectory(dir);
        using var c = new SqliteConnection($"Data Source={DbPath}");
        c.Open();
        var cmd = c.CreateCommand();
        cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS modules(
  id TEXT PRIMARY KEY, type TEXT, version TEXT, hosts TEXT, path TEXT, signed INTEGER
);
CREATE TABLE IF NOT EXISTS selections(
  profile_id TEXT, module_id TEXT, app TEXT, scope TEXT,
  PRIMARY KEY(profile_id, module_id, app)
);
CREATE TABLE IF NOT EXISTS ops(
  id TEXT PRIMARY KEY, ts_utc TEXT, action TEXT, app TEXT, module_id TEXT, scope TEXT, status TEXT, details TEXT
);
CREATE TABLE IF NOT EXISTS licenses(
  module_id TEXT PRIMARY KEY, kind TEXT, subject TEXT, expires_utc TEXT, raw BLOB
);
CREATE TABLE IF NOT EXISTS settings(
  k TEXT PRIMARY KEY, v TEXT
);
";
        cmd.ExecuteNonQuery();
    }
}
'@

# --- CLI Program ---
New-TextFile -Path src/Integrator.Cli/Program.cs -Content @'
using System.CommandLine;
using Integrator.Engine;

var root = new RootCommand("Company Integrator CLI");

var integrate = new Command("integrate", "Integrate add-ins");
var appsOpt = new Option<string[]>("--apps"){ Arity = ArgumentArity.OneOrMore, Description="Word,Excel,PowerPoint,Outlook" };
var scopeOpt = new Option<string>("--scope", () => "user"){ Description="user|machine" };
integrate.AddOption(appsOpt);
integrate.AddOption(scopeOpt);
integrate.SetHandler((string[] apps, string scope) =>
{
    var txn = new MutationTransaction(Paths.RollbackDir);
    try
    {
        foreach (var app in apps)
        {
            AddinRegistrarVstoCom.RegisterVsto(txn, scope.Equals("machine", StringComparison.OrdinalIgnoreCase),
                app, "Company.Reports", "Company Reports", "Reporting tools",
                Path.Combine(Paths.PerUserAddinsRoot, "Reports","Reports.vsto"));
        }
        txn.Commit();
        Console.WriteLine("Integrated successfully.");
    }
    catch (Exception ex)
    {
        txn.Rollback();
        Console.Error.WriteLine($"Integration failed: {ex.Message}");
        Environment.ExitCode = 2;
    }
}, appsOpt, scopeOpt);
root.Add(integrate);

var unintegrate = new Command("unintegrate", "Remove add-ins");
unintegrate.SetHandler(() =>
{
    var txn = new MutationTransaction(Paths.RollbackDir);
    try
    {
        foreach (var app in new[] {"Word","Excel","PowerPoint","Outlook"})
            AddinRegistrarVstoCom.Unregister(txn, false, app, "Company.Reports");
        txn.Commit();
        Console.WriteLine("Removed.");
    }
    catch (Exception ex)
    {
        txn.Rollback();
        Console.Error.WriteLine($"Unintegrate failed: {ex.Message}");
        Environment.ExitCode = 2;
    }
});
root.Add(unintegrate);

return await root.InvokeAsync(args);
'@

# --- Helper Admin stub ---
New-TextFile -Path src/Integrator.Helper.Admin/App.manifest -Content @'
<?xml version="1.0" encoding="utf-8"?>
<assembly manifestVersion="1.0" xmlns="urn:schemas-microsoft-com:asm.v1">
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security>
      <requestedPrivileges xmlns="urn:schemas-microsoft-com:asm.v3">
        <requestedExecutionLevel level="requireAdministrator" uiAccess="false" />
      </requestedPrivileges>
    </security>
  </trustInfo>
</assembly>
'@

New-TextFile -Path src/Integrator.Helper.Admin/Program.cs -Content @'
Console.WriteLine("Admin helper placeholder. Implement elevated operations as needed.");
'@

# --- Packaging project (MSIX) ---
New-TextFile -Path src/Integrator.Packaging/Integrator.Packaging.wapproj -Content @'
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <ItemGroup Label="ProjectConfigurations">
    <ProjectConfiguration Include="Debug|x64">
      <Configuration>Debug</Configuration><Platform>x64</Platform>
    </ProjectConfiguration>
    <ProjectConfiguration Include="Release|x64">
      <Configuration>Release</Configuration><Platform>x64</Platform>
    </ProjectConfiguration>
  </ItemGroup>
  <PropertyGroup>
    <ProjectGuid>{7F0F6C7E-7EAB-4E52-94E1-8D69B6B9C1A1}</ProjectGuid>
    <TargetPlatformVersion>10.0.19041.0</TargetPlatformVersion>
    <TargetPlatformMinVersion>10.0.19041.0</TargetPlatformMinVersion>
    <DefaultLanguage>en-US</DefaultLanguage>
    <PackageCertificateKeyFile></PackageCertificateKeyFile>
    <AppxBundle>Always</AppxBundle>
    <AppxBundlePlatforms>x64</AppxBundlePlatforms>
    <GenerateAppInstallerFile>False</GenerateAppInstallerFile>
  </PropertyGroup>
  <ItemGroup>
    <AppxManifest Include="Package.appxmanifest" />
  </ItemGroup>
  <ItemGroup>
    <None Include="Assets\StoreLogo.png" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\Integrator.UI.Wpf\Integrator.UI.Wpf.csproj" />
  </ItemGroup>
  <Import Project="$(WapProjPath)\Microsoft.DesktopBridge.targets" />
</Project>
'@

New-TextFile -Path src/Integrator.Packaging/Package.appxmanifest -Content @'
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
         xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
         IgnorableNamespaces="uap">
  <Identity Name="YourCompany.Integrator" Publisher="CN=Your Company Inc" Version="1.0.0.0"/>
  <Properties>
    <DisplayName>Company Office Integrator</DisplayName>
    <PublisherDisplayName>Your Company Inc</PublisherDisplayName>
    <Logo>Assets\StoreLogo.png</Logo>
  </Properties>
  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.19041.0" MaxVersionTested="10.0.26100.0" />
  </Dependencies>
  <Applications>
    <Application Id="App" Executable="Integrator.UI.Wpf.exe" EntryPoint="Windows.FullTrustApplication">
      <uap:VisualElements DisplayName="Company Office Integrator" Description="Integrates company add-ins with Office" />
    </Application>
  </Applications>
</Package>
'@

# dummy asset
New-TextFile -Path src/Integrator.Packaging/Assets/StoreLogo.png -Content ''

# --- Payload: Office Web Add-in manifest ---
New-TextFile -Path payload/AnalyticsManifest.xml -Content @'
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="TaskPaneApp">
  <Id>f9f9a3a4-0b8a-4a2f-9d1a-0b7b0b0c1d10</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Your Company Inc.</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Company Analytics"/>
  <Description DefaultValue="Task pane with analytics tools for Word &amp; Excel"/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://example.com/support"/>
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
'@

# --- Deploy & CI ---
New-TextFile -Path deploy/SmokeTest.ps1 -Content @'
$apps = "Word","Excel"
$addinId = "Company.Reports"
$catalog = "$env:ProgramData\Company\AddinCatalog"
$manifest = Join-Path $catalog "Company.Analytics.xml"

Write-Host "Validating registry for VSTO add-ins..." -ForegroundColor Cyan
$bad = $false
foreach ($app in $apps) {
  $k = "HKCU:\Software\Microsoft\Office\$app\Addins\$addinId"
  $ok = Test-Path $k
  $lb = if ($ok) { (Get-ItemProperty -Path $k -Name LoadBehavior -ErrorAction SilentlyContinue).LoadBehavior } else { $null }
  $mf = if ($ok) { (Get-ItemProperty -Path $k -Name Manifest -ErrorAction SilentlyContinue).Manifest } else { $null }
  if ($ok -and $lb -eq 3 -and $mf -and (Test-Path $mf)) { Write-Host "[$app] OK" -ForegroundColor Green }
  else { Write-Host "[$app] MISSING or INVALID" -ForegroundColor Red; $bad = $true }
}

Write-Host "Validating Trusted Catalog + Web manifest..." -ForegroundColor Cyan
$wef = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$tc = Get-ChildItem $wef -ErrorAction SilentlyContinue | Where-Object { (Get-ItemProperty $_.PSPath).Url -match [regex]::Escape($catalog) }
if ($tc -and (Test-Path $manifest)) { Write-Host "[Catalog] OK" -ForegroundColor Green } else { Write-Host "[Catalog] MISSING" -ForegroundColor Red; $bad = $true }

if ($bad) { Write-Host "Smoke test FAILED" -ForegroundColor Red; exit 1 } else { Write-Host "Smoke test PASSED" -ForegroundColor Green; exit 0 }
'@

New-TextFile -Path .github/workflows/build-release.yml -Content @'
name: Build & Release MSIX
on:
  push:
    tags: ["v*.*.*"]

jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4

      - name: Setup MSBuild
        uses: microsoft/setup-msbuild@v2

      - name: Setup NuGet
        uses: NuGet/setup-nuget@v2

      - name: Restore
        run: nuget restore DesktopIntegrator.sln

      - name: Build (Release x64)
        run: msbuild DesktopIntegrator.sln /p:Configuration=Release /p:Platform=x64

      - name: Build MSIX packaging project
        run: msbuild .\src\Integrator.Packaging\Integrator.Packaging.wapproj /p:Configuration=Release /p:Platform=x64 /t:Restore,Rebuild

      - name: Locate MSIX
        id: msix
        shell: pwsh
        run: |
          $msix = Get-ChildItem -Recurse -Filter *.msix | Sort-Object LastWriteTime -Descending | Select-Object -First 1
          if (-not $msix) { throw "MSIX not found" }
          echo "path=$($msix.FullName)" >> $env:GITHUB_OUTPUT

      - name: Create GitHub Release
        uses: softprops/action-gh-release@v2
        with:
          files: ${{ steps.msix.outputs.path }}
          name: Release ${{ github.ref_name }}
          draft: false
          prerelease: false
'@

New-TextFile -Path docs/index.html -Content @'
<!DOCTYPE html><html><head><meta charset="utf-8"><title>Company Office Integrator</title></head>
<body>
  <h1>Company Office Integrator</h1>
  <p>Download the latest installer:</p>
  <p><a href="https://github.com/<your-org>/desktop-integrator/releases/latest/download/CompanyOfficeIntegrator.msix">
    CompanyOfficeIntegrator.msix
  </a></p>
  <p>Replace &lt;your-org&gt; with your GitHub org or username.</p>
</body></html>
'@

# .gitignore for .NET/VS
New-TextFile -Path .gitignore -Content @'
## Ignore Visual Studio and build artifacts
.vs/
bin/
obj/
*.user
*.suo
*.opendb
*.db
*.pfx
*.snk
*.msix
*.msixbundle
*.appx
*.appxbundle
'@

# README (short)
New-TextFile -Path README.md -Content @'
# Company Office Integrator (Windows 11 Pro)

## Build
- Open `DesktopIntegrator.sln` in Visual Studio 2022
- Set `Integrator.UI.Wpf` as start-up project
- Run (F5)

## Package (MSIX)
- Build `src/Integrator.Packaging/Integrator.Packaging.wapproj` (Release x64)

## Release
- Tag a version: `git tag v1.0.0 && git push origin v1.0.0`
- Download (latest): `https://github.com/<your-org>/desktop-integrator/releases/latest/download/CompanyOfficeIntegrator.msix`

## Notes
- Web add-in manifest is copied to `%ProgramData%\Company\AddinCatalog\Company.Analytics.xml`
- Sample VSTO manifest expected at `%LOCALAPPDATA%\Company\Addins\Reports\Reports.vsto` (a placeholder is created automatically for demo)
'@

Write-Host "==> Restoring packages..." -ForegroundColor Cyan
& dotnet restore DesktopIntegrator.sln | Out-Null

Write-Host "==> Formatting done. Commit these files to your repo." -ForegroundColor Green
Write-Host "   Next steps:"
Write-Host "     1) git add . && git commit -m 'feat: initial integrator repo'"
Write-Host "     2) git remote add origin https://github.com/<your-org>/desktop-integrator.git"
Write-Host "     3) git push -u origin main"
Write-Host "     4) git tag v1.0.0 && git push origin v1.0.0"
