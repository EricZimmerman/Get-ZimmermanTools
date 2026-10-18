# Get-ZimmermanTools
This script will auto-discover all available downloads and download what does not already exist in a given destination directory.

The script does not check for what is on disk, but uses what is in a CSV file that this script generates to know what to download based on SHA-1 hash.

# Build-ZimmermanTools (maintainers)

`Build-ZimmermanTools.ps1` builds every EZ Tools program from its repository (cloned next to this one), stages each as its download zip is laid out, signs them and makes the download zips. It is for building releases; users only need `Get-ZimmermanTools.ps1`.

```powershell
.\Build-ZimmermanTools.ps1 --out D:\Release                              # build and stage everything
.\Build-ZimmermanTools.ps1 --out D:\Release --only RECmd,TimelineExplorer
.\Build-ZimmermanTools.ps1 --sign D:\Release                             # signtool over every .exe/.dll (password asked once per call)
.\Build-ZimmermanTools.ps1 --staged D:\Release --zip \server\share      # the download zips (.NET Framework at the top, .NET 10 in net\)
.\Build-ZimmermanTools.ps1 --out D:\Release --sign --zip \server\share  # all three
```

`Get-Help .\Build-ZimmermanTools.ps1 -Full` has the details. It needs the .NET 10 SDK, Visual Studio's MSBuild for old-style projects, the DevExpress NuGet feed for the programs with a window, and signtool for `--sign`.
