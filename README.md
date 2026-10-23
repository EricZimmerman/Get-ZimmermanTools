# Get-ZimmermanTools
This script will auto-discover all available downloads and download what does not already exist in a given destination directory.

The script does not check for what is on disk, but uses what is in a CSV file that this script generates to know what to download based on SHA-1 hash.

Before downloading anything, the script checks for a newer version of itself. If there is one, it replaces the script with it and stops: run it again to use the new version. `-NoUpdate` skips the check.

# Build-ZimmermanTools (maintainers)

`Build-ZimmermanTools.ps1` builds every EZ Tools program from its repository (cloned next to this one), stages each as its download zip is laid out, signs them and makes the download zips. It is for building releases; users only need `Get-ZimmermanTools.ps1`.

```powershell
.\Build-ZimmermanTools.ps1 --out D:\Release                              # build and stage everything
.\Build-ZimmermanTools.ps1 --out D:\Release --tool LECmd,MFTECmd            # only these (also -Tool / -Only)
.\Build-ZimmermanTools.ps1 --out D:\Release --clear                       # empty D:\Release first
.\Build-ZimmermanTools.ps1 --sign D:\Release                             # signtool over the files _sign.txt lists (password asked once per call)
.\Build-ZimmermanTools.ps1 --staged D:\Release --zip \server\share      # the download zips (.NET Framework at the top, .NET 10 in net\)
.\Build-ZimmermanTools.ps1 --out D:\Release --sign --zip \server\share  # all three
```

`Get-Help .\Build-ZimmermanTools.ps1 -Full` has the details. It needs the .NET 10 SDK, Visual Studio's MSBuild for old-style projects, the DevExpress NuGet feed for the programs with a window, and signtool for `--sign`.

The command line tools carry their dependencies inside (Costura) and stay cross platform (`dotnet <Tool>.dll`). Native libraries are files next to them, never inside them or extracted when they run: AmcacheParser and WxTCmd's SQLite (`e_sqlite3.dll`, `libe_sqlite3.so`, `libe_sqlite3.dylib`: Windows x64, Linux x64, macOS arm64), SQLECmd's (`runtimes\<rid>\native\SQLite.Interop.dll`: Windows x64/x86, Linux x64, macOS x64). Staging fails if a tool has a native library inside, and staging and `--zip` fail if two zips that Get-ZimmermanTools extracts into the same folder have different files at the same path (AmcacheParser and WxTCmd share their SQLite files at the top of `net\`). The programs with a window are each one exe (.NET single file: framework dependent, win-x64, published with `-p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=false`), staged and zipped in its own folder with what the publish leaves next to it: native DLLs such as `sni.dll` (never extracted to `%TEMP%` when the program runs), `<Program>.dll.config` and content files, but no pdbs. Staging reads the exe's bundle manifest and fails if a file its `deps.json` lists is neither inside the exe nor next to it, if a managed DLL is loose next to it (plugins in `Plugins` excepted), or if a file the publish made is not staged.

No program ships a `.pdb`: the pdbs are embedded in the assemblies (each repository's `Directory.Build.props`), so stack traces and error reports have file and line numbers. Staging fails a program if a `.pdb` ends up staged, and warns (without failing) if the program's own assembly has no embedded pdb or one with this machine's source paths (`D:\Code\...`) instead of `/_/...` (a command line tool's `<Tool>.dll`, or `<Tool>.exe` for .NET Framework; a single file's `<Program>.dll`, read from its bundle).

`--sign` signs only Eric's files: staging writes `_sign.txt` at the top of the staged folder (every `.exe` and `.ps1` (a single-file exe is signed whole, with its bundle), each program's own dll, the plugins, and the dlls of his projects and NuGet packages, by the `.nuspec` authors in the NuGet cache). Third-party dlls (Serilog, CsvHelper, ...), Microsoft's and DevExpress' are left unsigned, and are never signed even with `--force`. A staged folder without `_sign.txt` (staged by an older version) is not signed: restage it. `_sign.txt` is in no zip.
