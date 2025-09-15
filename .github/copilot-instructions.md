# Copilot Instructions for Company Office Integrator

## Project Overview
This is a Windows desktop integration suite for managing Office add-ins (VSTO, Web) and related automation. The solution is multi-project, built with .NET 8, and packaged as an MSIX installer. The main UI is a WPF app (`src/Integrator.UI.Wpf`), with supporting engine, shared, CLI, and admin helper projects.

## Architecture & Components
- **WPF UI (`src/Integrator.UI.Wpf`)**: Main user interface, MVVM pattern, binds to `IntegrateViewModel`. Uses `RelayCommand` for actions.
- **Engine (`src/Integrator.Engine`)**: Handles registry/file mutations via `MutationTransaction`, add-in registration (`AddinRegistrarVstoCom`, `AddinRegistrarWeb`), and path utilities (`Paths`).
- **Shared (`src/Integrator.Shared`)**: Logging (`LogInit`), SQLite storage (`Storage`), and contracts.
- **CLI (`src/Integrator.Cli`)**: Command-line integration/removal for add-ins. Uses `System.CommandLine`.
- **Admin Helper (`src/Integrator.Helper.Admin`)**: Placeholder for elevated operations.
- **Packaging (`src/Integrator.Packaging`)**: MSIX packaging project (`.wapproj`).
- **Payload**: Contains Office Web Add-in manifest XML.
- **Deploy**: Contains PowerShell smoke test for registry/catalog validation.

## Build & Run
- Open `DesktopIntegrator.sln` in Visual Studio 2022.
- Set `Integrator.UI.Wpf` as startup project and run (F5).
- Restore packages: `dotnet restore DesktopIntegrator.sln`.
- Build MSIX: Build `src/Integrator.Packaging/Integrator.Packaging.wapproj` (Release x64).
- CLI usage: `dotnet run --project src/Integrator.Cli -- integrate --apps Word Excel --scope user`

## Key Patterns & Conventions
- **Registry and file changes** are always wrapped in `MutationTransaction` for rollback support.
- **Add-in registration**: Use `AddinRegistrarVstoCom.RegisterVsto` for VSTO, `AddinRegistrarWeb` for web add-ins.
- **MVVM**: UI logic is in `IntegrateViewModel`, with observable properties and commands.
- **Smoke tests**: Use `deploy/SmokeTest.ps1` to validate registry/catalog state after integration.
- **Logging**: Use `LogInit.Init(logDir)` to initialize Serilog file logging.
- **SQLite**: Use `Storage.Ensure()` to initialize the local DB.

## External Dependencies
- .NET 8 SDK (Windows)
- NuGet packages: `Microsoft.Data.Sqlite`, `Serilog`, `System.CommandLine`, `Microsoft.Identity.Client`
- MSIX packaging tools for deployment

## Integration Points
- Registry: Office add-in registration under `HKCU/HKLM\Software\Microsoft\Office\<App>\Addins`
- File system: Manifest files in `%LOCALAPPDATA%` and `%ProgramData%`
- Web add-in catalog: `%ProgramData%\Company\AddinCatalog`

## Example Workflow
1. User selects apps in UI, clicks "Run" to register add-ins (registry + manifest).
2. "Rollback" undoes changes via transaction log.
3. CLI and smoke test scripts provide automation and validation.

## References
- See `README.md` for build/package/release steps.
- See `deploy/SmokeTest.ps1` for validation logic.
- See `src/Integrator.UI.Wpf/IntegrateViewModel.cs` for main integration logic.

---
If any section is unclear or missing details, please specify which workflows, patterns, or files need further explanation.