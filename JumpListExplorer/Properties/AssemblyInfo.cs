using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

[assembly: AssemblyCopyright("Copyright (C) 2022-2025 Simon Mourier. All rights reserved.")]
[assembly: AssemblyTitle("Windows Jump List Explorer")]
#if DEBUG
[assembly: AssemblyConfiguration("DEBUG")]
#else
[assembly: AssemblyConfiguration("RELEASE")]
#endif
[assembly: AssemblyDescription("Windows Jump List Explorer")]
[assembly: AssemblyCompany("Simon Mourier")]
[assembly: AssemblyProduct("Windows Jump List Explorer")]
[assembly: AssemblyCulture("")]

[assembly: ComVisible(false)]
[assembly: Guid("c9ce0ba6-4c4e-4d31-b1c2-e002a78f83fc")]
[assembly: SupportedOSPlatform("windows")]