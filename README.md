# OutlookMailViewerMVVM

A minimal WPF (.NET 8) MVVM app that connects to Outlook (desktop) and lists Inbox emails
between two dates into a DataGrid: **Received**, **Subject**, **Sender**.

## Requirements
- Windows + Outlook desktop installed and configured with a profile.
- .NET 8 SDK.
- Visual Studio 2022 or newer.

## Build & Run
1. Open the solution folder in VS and restore packages.
2. Build and run `OutlookMailViewerMVVM` (x64).

## Notes
- The app uses `Microsoft.Office.Interop.Outlook` (NuGet). Outlook must be installed.
- Dates in the Outlook Restrict filter are formatted with `en-US` culture as required by the interop.
- COM objects are explicitly released to avoid memory leaks.


## Runtime note
This project uses COM references with **EmbedInteropTypes=true** for Outlook/Office. It does not require
the 'office' PIA assembly at runtime. Windows + Outlook desktop must be installed (same bitness as build, x64).
If you still get COM activation errors, repair Office installation and ensure Outlook can be opened manually.
