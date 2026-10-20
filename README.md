# Get-ZimmermanTools
This script will auto-discover all available downloads and download what does not already exist in a given destination directory.

The script does not check for what is on disk, but uses what is in a CSV file that this script generates to know what to download based on SHA-1 hash.

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

The command line tools carry their dependencies inside (Costura); the programs with a window don't (they start several times faster without it): each is staged and zipped in its own folder with its DLLs and `runtimes\` next to the exe, and staging fails if a file their `.deps.json` lists is missing.

`--sign` signs only Eric's files: staging writes `_sign.txt` at the top of the staged folder (every `.exe` and `.ps1`, each program's own dll, the plugins, and the dlls of his projects and NuGet packages, by the `.nuspec` authors in the NuGet cache). Third-party dlls (Serilog, CsvHelper, ...), Microsoft's and DevExpress' are left unsigned, and are never signed even with `--force`. A staged folder without `_sign.txt` (staged by an older version) is not signed: restage it. `_sign.txt` is in no zip.
