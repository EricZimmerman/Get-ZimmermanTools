# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Two PowerShell scripts:

- **`Get-ZimmermanTools.ps1`** (for users; ships as `Get-ZimmermanTools.zip`): finds every download on https://tools.ericzimmermanstools.com and downloads what changed (by SHA-1, kept in a CSV) to `-Dest`: .NET Framework zips from the site's top level, .NET 10 ones from `/net/` into `<Dest>\net` (`-NetVersion 0|4|10`, default 10; 9 maps to 10 with a warning). Must keep working in Windows PowerShell 5.1 and PowerShell 7.
- **`Build-ZimmermanTools.ps1`** (for the maintainer; not in the user zip): builds every EZ Tools program from its repository, cloned next to this one (`-Root`, default the parent folder), stages each as its download zip is laid out (`-Out`), signs (`-Sign`), and makes the zips (`-Zip`: .NET Framework at the top, .NET 10 in `net\`). `-Tool a,b` (`-Only`) builds a few, `-Clear` empties `-Out` first (never a drive, a repository or the repositories' folder). GNU style switches (`--out`, `--tool`, ...) work too. `Get-Help .\Build-ZimmermanTools.ps1 -Full` has the rest.

## Build-ZimmermanTools: how it works

- **The program table** (`$programs`) is the description of every download zip: `Project`, `Folder` (the folder in the zip), `Files` (from the build besides the defaults), `Copy` (repository folders: EvtxECmd and SQLECmd maps, RECmd batch files, Registry Explorer bookmarks and manual), `Release` (files only in the current release zips, downloaded once into `%TEMP%\Build-ZimmermanTools\_release`: LICENSE.txt, GeoLite2-City.mmdb, XWFIM's text files, Registry Explorer's default settings), `Special` (the registry and TLE plugins, built with each repository's `CopyPlugins.ps1 -Build`; TLE's layouts come from TLEFilePlugins `Layouts\Default`, their source of truth), `SingleFile` (AmcacheParser and WxTCmd: published as one exe, framework dependent, win-x64, native SQLite inside; they have no Costura). A tool whose output or content changes needs its entry changed here. **ezTriage is not in the table and must not be added** (Eric, 2026-10-20): it has its own release script, `Build-Release.ps1` in the ezTriage repository.
- **Content comes from the local clones**, not GitHub; the maintainer keeps them in sync with the remotes. The tools' own `--sync` downloads the same folders from GitHub `master` and compares bytes, so the `Copy` folders and TLE's layouts are taken as committed at HEAD (`Export-RepoPath`: `git archive` with `core.autocrlf=false`), not from the working copy, whose CRLF made `--sync` report every map as updated. Uncommitted changes there are left out with a warning. The script itself (`Get-ZimmermanTools.ps1`) stays the working copy's CRLF, as released before.
- **Builds**: `dotnet build -f <tfm> -o %TEMP%\Build-ZimmermanTools\_build\<Program>\<fw> --no-incremental` (local, never in `-Out`, which is often a share; `--clear` empties it too) (full builds: an incremental build skipped Costura weaving once). Visual Studio's MSBuild only for old-style projects or a `licenses.licx` embedded for any MSBuild; MSBuild 18 ends with 0x4000001E when given `OutDir`/`TargetFramework` or with dotnet build server nodes left, so the script restores and builds separately and shuts the nodes down.
- **Checks**: every dll the build made must be staged (anywhere in the program) or embedded by Costura (read from the assembly's metadata); a Costura program with nothing embedded fails. Failures are listed in the summary table and set the exit code.
- **The script**: the `Get-ZimmermanTools` entry (`Script`) copies `Get-ZimmermanTools.ps1` into `<Out>\Get-ZimmermanTools\script`; `--sign` signs that copy (the repository's stays unsigned) and `--zip` makes `Get-ZimmermanTools.zip` at the top.
- **Two machines**: Eric stages on the dev machine and signs on another (hardware token). Staging copies this script and writes `_staged.txt` at the top of `-Out`; that copy, run with `--sign --zip <dir>` and no `--out`/`--staged`, takes its own folder as the staged one. Sign and zip must keep working without repositories and in Windows PowerShell 5.1 (no `$PSScriptRoot` in param defaults, no `File.Move` overwrite, `[NullString]::Value` for a null string argument).
- **Signing**: every `.exe`, `.dll` and `.ps1` in the program folders (not the script copy at the top); `signtool sign /fd sha256 /tr <timestamp> /td sha256` (`/a`, or `/f` with `-CertFile`), many files per call so a token asks once per call; files already validly signed are skipped unless `-Force`. Never run signing on a machine without the user asking: `/a` picks the user's real certificate.
- **Zips** (`--zip`): made locally, copied next to the destination zip as `<name>.zip.partial-<guid>`, then renamed over it (one step, also on a share): a download meanwhile gets the whole old zip. Nothing is deleted at the destination (an old `net9\` folder or a zip no longer made stays).
- Output: each program's `== Program fw` line in red, helper scripts' messages indented under it, then a summary table.

## Testing

No automated tests. Verify a change by staging a few programs to a temp folder (`--out <temp> --tool A,B --clear`) and comparing each staged folder with the current release zip (same files, except intended changes) and with what the build produced. A full run (`--out`) builds everything and takes a while.

## Versioning and release

No version. `Get-ZimmermanTools.ps1` is released by zipping it to the download share (`Get-ZimmermanTools.zip`, the script's `script` entry). The website (`ericzimmerman.github.io`) links to the same URLs: its links and the share's layout must change together.
