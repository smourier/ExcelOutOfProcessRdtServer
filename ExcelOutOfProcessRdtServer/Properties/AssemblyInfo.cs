using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Windows;

#if DEBUG
[assembly: AssemblyConfiguration("DEBUG")]
#else
[assembly: AssemblyConfiguration("RELEASE")]
#endif

[assembly: AssemblyTitle("ExcelOutOfProcessRdtServer")]
[assembly: AssemblyProduct("ExcelOutOfProcessRdtServer")]
[assembly: AssemblyCopyright("Copyright (C) 2023-2024 Simon Mourier. All rights reserved.")]
[assembly: AssemblyCulture("")]
[assembly: AssemblyDescription("An Excel out-of-process Real-Rime Data (RDT) server written in .NET Core.")]
[assembly: AssemblyCompany("Simon Mourier")]
[assembly: Guid("d802dfc2-80d2-4c54-b8f4-393bdc107ace")]
[assembly: ThemeInfo(ResourceDictionaryLocation.None, ResourceDictionaryLocation.SourceAssembly)]
[assembly: SupportedOSPlatform("windows")]
