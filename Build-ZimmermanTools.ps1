<#
.SYNOPSIS
    Builds every EZ Tools program in Release, stages each one as its download zip is laid out, and signs them.

.DESCRIPTION
    Staging (-Out, or --out): every program is built from its repository (under -Root) for each framework its project
    targets, and staged in <Out>\<Program>\<framework> (net4, net10, ...), laid out as in its download zip: the
    single-file command line tools as their files, the others in a folder (RECmd\RECmd\..., RegistryExplorer\RegistryExplorer\...).

      - EvtxECmd, SQLECmd: their Maps; RECmd: BatchExamples and the registry plugins (Plugins)
      - Registry Explorer: the registry plugins, the bookmarks (RegistryExplorerBookmarks\Common), the manual
      - Timeline Explorer: the plugin DLLs (Plugins) and the layouts (Layouts\Default) from TLEFilePlugins
      - The programs with a window are one exe each (.NET single file: framework dependent, win-x64; their managed
        dependencies are inside it), in their own folder with what the publish leaves next to it: native dlls such as
        sni.dll (never extracted to %TEMP%), <Program>.dll.config, content files (EZViewer's Extensions.yaml). No pdbs.
        The command line tools stay cross platform: <Tool>.dll with its dependencies inside (Costura), run with
        <Tool>.exe or dotnet <Tool>.dll; AmcacheParser and WxTCmd are single files (win-x64)
      - Maps, batch files, bookmarks, the manual and the layouts are taken as committed (HEAD), byte for byte what
        GitHub has (and what --sync downloads), not from the working copy (CRLF on Windows); uncommitted changes in
        them are left out with a warning
      - Files that are not in any repository (the GUIs' LICENSE.txt, iisGeolocate's GeoLite2-City.mmdb, Registry
        Explorer's default Settings, XWFIM's text files) are taken from the current release zips, downloaded once
        (from the page Get-ZimmermanTools.ps1 reads) into %TEMP%\Build-ZimmermanTools\_release

    The build output is in %TEMP%\Build-ZimmermanTools\_build: -Out (which can be a share) only gets the staged
    programs, a copy of this script, _staged.txt and _sign.txt. The copy is replaced by every staging run, so the one in -Out is
    always the current one.

    Command line tools are built with dotnet build, programs with a window published with dotnet publish
    (-r win-x64 --self-contained false -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=false); Visual
    Studio's MSBuild only builds old-style projects and those that embed licenses.licx for it.

    Checks: every dll the build made is staged or inside the program (Costura's resources, or a single file's bundle,
    read from its manifest); a program with a window has every file its bundled deps.json lists inside the exe or next to
    it, no managed dll loose next to it (they belong in the exe; plugins in Plugins excepted), and every file the publish
    made (but pdbs) staged (next to the exe as published; Special may replace those in Plugins and Layouts).

    Get-ZimmermanTools.ps1 is staged too (<Out>\Get-ZimmermanTools\script), so it is signed and zipped with the rest;
    -Only Get-ZimmermanTools stages just the script.

    Staging also writes _sign.txt at the top of -Out: the files to sign, one path per line. Only Eric's: every .exe and
    .ps1 of the programs (a single-file exe is signed as a whole, its bundle with it), each program's own dll, the
    registry and Timeline Explorer plugins, and the dlls of his projects and NuGet packages (a deps.json "project", or a
    package whose .nuspec in the NuGet cache names Eric Zimmerman). Third-party dlls (Serilog, CsvHelper, ...),
    Microsoft's (sni.dll) and DevExpress' are not listed. It is decided
    when staging, since the signing machine has no build output or NuGet cache; it is in no zip.

    Signing (-Sign <dir>, or --sign): signs the files _sign.txt lists with signtool.exe
    (sign /fd sha256 /tr <timestamp> /td sha256), many files per call, so a token asks for its password once per call:
    with the certificate picked automatically (/a, the default), or from a .pfx (-CertFile; its password is asked once).
    Files that already have a valid signature are left alone unless -Force is used (which re-signs only listed files).
    Microsoft's and DevExpress' files are never signed, even listed or with -Force. A listed file that is missing is
    reported (and the zips are skipped). A staged folder without _sign.txt (staged by an older version) is not signed:
    restage it.

    Zips (-Zip <dir>, or --zip): one zip per program and framework from the staged folder (-Out, or -Staged for a
    folder staged and signed earlier), laid out as the download folder: the .NET Framework builds (and
    Get-ZimmermanTools.zip) at the top, the others in net\ (Get-ZimmermanTools.ps1 looks for /net/).

    They can be combined; they run in this order: build and stage, sign, zip (so the zips have the signed files).
    -Sign without a folder signs the staged folder. With -Sign, a signing failure skips the zips.

    Signing on another machine (a hardware token): staging leaves a copy of this script and _staged.txt (what was
    staged, any failures, the commands) at the top of -Out. On the other machine, run that copy with --sign --zip
    <download folder>: it takes its own folder as the staged one and needs no repositories (Windows PowerShell 5.1
    or PowerShell 7).

.PARAMETER Out
    Where to stage the programs. Also --out.

.PARAMETER Sign
    The staged folder to sign (the files its _sign.txt lists). Also --sign. Without a value, the -Out or -Staged folder.

.PARAMETER Root
    The folder that has the repositories (default: the folder above this repository, where the other repositories are
    cloned next to it).

.PARAMETER Only
    Build only these programs, comma separated (names as in the zips: LECmd,MFTECmd,TimelineExplorer, ...). Also -Tool,
    --only a,b and --tool a,b.

.PARAMETER Clear
    Delete everything in the -Out folder first, and the build output and cached release zips (in %TEMP%), so
    nothing is left from an earlier run. Also --clear.

.PARAMETER Zip
    Make the download zips in this folder (<Zip>\X.zip for .NET Framework, <Zip>\net\X.zip for .NET 10). Also --zip.

.PARAMETER Staged
    A folder staged earlier (by -Out), to sign or zip without building. Also --staged.

.PARAMETER CertFile
    Sign with this .pfx instead of a certificate from the store; its password is asked once.

.PARAMETER Force
    Sign the listed files that already have a valid signature too (never Microsoft's or DevExpress').

.PARAMETER SignTool
    The signtool.exe to use (default: on the PATH, else the newest Windows SDK's, the App Certification Kit's or
    Visual Studio's ClickOnce one). Also --signtool.

.PARAMETER TimestampUrl
    The RFC 3161 time stamp server (default http://timestamp.sectigo.com).

.EXAMPLE
    .\Build-ZimmermanTools.ps1 --out D:\Release
    .\Build-ZimmermanTools.ps1 --sign D:\Release
    .\Build-ZimmermanTools.ps1 --staged D:\Release --zip \\server\share
    .\Build-ZimmermanTools.ps1 --out D:\Release --sign --zip \\server\share
    .\Build-ZimmermanTools.ps1 -Out D:\Release -Only RECmd,TimelineExplorer
    .\Build-ZimmermanTools.ps1 -Out D:\Release -Clear -Tool LECmd,MFTECmd
    \\server\stage\Build-ZimmermanTools.ps1 --sign --zip \\server\share    (on the signing machine, after --out \\server\stage)
#>
[CmdletBinding(PositionalBinding = $false)]
param(
    [string] $Out,
    [string] $Sign,
    [string] $Root,
    [Alias('Tool')]
    [string[]] $Only,
    [switch] $Clear,
    [string] $Zip,
    [string] $Staged,
    [string] $CertFile,
    [switch] $Force,
    [string] $TimestampUrl = 'http://timestamp.sectigo.com',
    [string] $SignTool,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $Rest
)

$ErrorActionPreference = 'Stop'
#(here, not as the parameter's default: Windows PowerShell 5.1 has no $PSScriptRoot there)
if (-not $Root) { $Root = Split-Path $PSScriptRoot -Parent }

# the build output (_build) and the cached release zips (_release): on this machine, so -Out (a share) only gets what is
# staged
$Work = Join-Path ([IO.Path]::GetTempPath()) 'Build-ZimmermanTools'

# --out x / --sign x / --only a,b (--tool a,b) / --clear / --zip / --force / --certfile x (GNU style)
for ($i = 0; $i -lt $Rest.Count; $i++) {
    switch -Regex ($Rest[$i]) {
        '^--?out$' { $Out = $Rest[++$i] }
        '^--?sign$' {
            if ($i + 1 -lt $Rest.Count -and $Rest[$i + 1] -notmatch '^--?\w') { $Sign = $Rest[++$i] } else { $Sign = '<out>' }
        }
        #(PowerShell passes a,b as an array, which arrives joined with spaces)
        '^--?(only|tool)$' { $Only = $Rest[++$i] -split '[,\s]+' | Where-Object { $_ } }
        '^--?clear$' { $Clear = $true }
        '^--?zip$' { $Zip = $Rest[++$i] }
        '^--?staged$' { $Staged = $Rest[++$i] }
        '^--?force$' { $Force = $true }
        '^--?certfile$' { $CertFile = $Rest[++$i] }
        '^--?signtool$' { $SignTool = $Rest[++$i] }
        '^--?root$' { $Root = $Rest[++$i] }
        default { throw "Unknown argument '$($Rest[$i])'. Use --out <dir>, --clear, --tool a,b, --sign [<dir>], --zip <dir>, --staged <dir>, --certfile <pfx>, --force" }
    }
}

# the copy that staging leaves in the staged folder (for signing on another machine): that folder is the staged one
$stagedFrom = $null
if (-not $Out -and -not $Staged -and ($Sign -or $Zip) -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot '_staged.txt'))) {
    $Staged = $PSScriptRoot
    $stagedFrom = Join-Path $PSScriptRoot '_staged.txt'
}

if (-not $Out -and -not $Sign -and -not $Zip) {
    throw "Nothing to do: give --out <dir> to build and stage, --sign <dir> to sign, --zip <dir> (with --out or --staged) to make the zips"
}
if ($Sign -eq '<out>') {
    $Sign = if ($Out) { $Out } elseif ($Staged) { $Staged } else { throw "--sign without a folder needs --out or --staged" }
}
if ($Zip -and -not $Out -and -not $Staged) { throw "--zip needs the staged folder: --out (build first) or --staged" }

#("LECmd,MFTECmd" in quotes arrives as one name)
if ($Only) { $Only = @($Only -split '[,\s]+' | Where-Object { $_ }) }

if ($Clear) {
    if (-not $Out) { throw "--clear needs --out (the folder to empty)" }
    $outFull = [IO.Path]::GetFullPath($Out).TrimEnd('\')
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')

    # never a drive, a folder with the repositories (or above them), a repository, or a profile or system folder
    $refuse = $outFull -eq [IO.Path]::GetPathRoot($outFull).TrimEnd('\') -or
        "$rootFull\".StartsWith("$outFull\", [StringComparison]::OrdinalIgnoreCase) -or
        (Test-Path -LiteralPath (Join-Path $outFull '.git')) -or
        [bool](@($env:USERPROFILE, $env:SystemRoot, $env:ProgramFiles, ${env:ProgramFiles(x86)}, [Environment]::GetFolderPath('Desktop'),
            [Environment]::GetFolderPath('MyDocuments')) | Where-Object { $_ -and $_.TrimEnd('\') -ieq $outFull })
    if ($refuse) { throw "--clear will not empty '$outFull' (a drive, a repository, the folder with the repositories or a system folder)" }

    if (Test-Path -LiteralPath $outFull) {
        $items = @(Get-ChildItem -LiteralPath $outFull -Force)
        Write-Host "Clearing $outFull ($($items.Count) item(s))" -ForegroundColor Yellow
        $items | Remove-Item -Recurse -Force
    }
    if (Test-Path -LiteralPath $Work) {
        Write-Host "Clearing $Work (the build output and the cached release zips)" -ForegroundColor Yellow
        Remove-Item -LiteralPath $Work -Recurse -Force
    }
}

# ---------------------------------------------------------------------------------------------------------------------
# What each download zip has. Project: under -Root. Folder: the folder in the zip ($null: the files at the top).
# Files: what is taken from the build output ({0} is the assembly name) besides the defaults (net4: {0}.exe;
# others: {0}.exe, {0}.dll, {0}.runtimeconfig.json). Gui with a Folder (.NET, not net4): published as one exe
# (Test-SingleFileGui: framework dependent, win-x64, native libraries next to it) and staged as the publish output
# without pdbs; Files are then only checked to be there. Copy: repository folders/files -> paths in the staged program.
# Release: files only in the release zips. Special: steps for plugins and layouts. SingleFile: published as one exe
# (framework dependent, win-x64, native libraries inside: the command line tools with SQLite, which have no Costura).
# ---------------------------------------------------------------------------------------------------------------------
$programs = @(
    @{ Name = 'AmcacheParser'; Project = 'AmcacheParser\AmcacheParser\AmcacheParser.csproj'; SingleFile = $true }
    @{ Name = 'AppCompatCacheParser'; Project = 'AppCompatCacheParser\AppCompatCacheParser\AppCompatCacheParser.csproj' }
    @{ Name = 'bstrings'; Project = 'bstrings\bstrings\bstrings.csproj' }
    @{ Name = 'EvtxECmd'; Project = 'evtx\EvtxECmd\EvtxECmd.csproj'; Folder = 'EvtxeCmd'; Copy = @{ 'evtx\evtx\Maps' = 'Maps' } }
    @{ Name = 'iisGeolocate'; Project = 'iisgeolocate\iisGeolocate\iisGeolocate.csproj'; Folder = 'iisGeolocate'; Release = @('GeoLite2-City.mmdb') }
    @{ Name = 'JLECmd'; Project = 'JLECmd\JLECmd\JLECmd.csproj' }
    @{ Name = 'LECmd'; Project = 'LECmd\LECmd\LECmd.csproj' }
    @{ Name = 'MFTECmd'; Project = 'MFTECmd\MFTECmd\MFTECmd.csproj' }
    @{ Name = 'PECmd'; Project = 'PECmd\PECmd\PECmd.csproj' }
    @{ Name = 'RBCmd'; Project = 'RBCmd\RBCmd\RBCmd.csproj' }
    @{ Name = 'RecentFileCacheParser'; Project = 'RecentFileCacheParser\RecentFileCacheParser\RecentFileCacheParser.csproj' }
    @{ Name = 'RECmd'; Project = 'RECmd\RECmd\RECmd.csproj'; Folder = 'RECmd'; Copy = @{ 'RECmd\BatchExamples' = 'BatchExamples' }; Special = 'RegistryPlugins' }
    @{ Name = 'rla'; Project = 'RECmd\rla\rla.csproj' }
    @{ Name = 'SBECmd'; Project = 'ShellBagsExplorer\SBE\SBE.csproj' }
    @{ Name = 'SQLECmd'; Project = 'SQLECmd\SQLECmd\SQLECmd.csproj'; Folder = 'SQLECmd'; Copy = @{ 'SQLECmd\SQLMap\Maps' = 'Maps' } }
    @{ Name = 'SrumECmd'; Project = 'srum\SrumECmd\SrumECmd.csproj' }
    @{ Name = 'SumECmd'; Project = 'sum\SumECmd\SumECmd.csproj' }
    @{ Name = 'VSCMount'; Project = 'vscmount\VSCMount\VSCMount.csproj' }
    @{ Name = 'WxTCmd'; Project = 'WxTCmd\WxTCmd\WxTCmd.csproj'; SingleFile = $true }

    @{ Name = 'EZViewer'; Project = 'EZViewer\EZViewer\EZViewer.csproj'; Folder = 'EZViewer'; Gui = $true; Files = @('Extensions.yaml'); Release = @('LICENSE.txt') }
    @{ Name = 'Hasher'; Project = 'Hasher\Hasher2\Hasher2.vbproj'; Folder = 'Hasher'; Gui = $true; Release = @('LICENSE.txt'); ZipName = 'hasher' }
    @{ Name = 'JumpListExplorer'; Project = 'jumplistexplorer\JumpListExplorer\JumpListExplorer.csproj'; Folder = 'JumpListExplorer'; Gui = $true; Release = @('LICENSE.txt') }
    @{ Name = 'MFTExplorer'; Project = 'MFTExplorer\MFTExplorer\MFTExplorer.csproj'; Folder = 'MFTExplorer'; Gui = $true; Release = @('LICENSE.txt') }
    @{ Name = 'RegistryExplorer'; Project = 'RegistryExplorer\RegistryViewerZ\RegistryViewerZ.csproj'; Folder = 'RegistryExplorer'; Gui = $true
        Copy = @{ 'RegistryExplorer\RegistryViewerZ\RegistryExplorerManual.pdf' = 'RegistryExplorerManual.pdf'; 'RegistryExplorerBookmarks\Common' = 'Bookmarks\Common' }
        Release = @('LICENSE.txt', 'Settings\*'); Special = 'RegistryPlugins'; EmptyDirs = @('Bookmarks\User') }
    @{ Name = 'SDBExplorer'; Project = 'sdbexplorer\SDBExplorer\SDBExplorer.csproj'; Folder = 'SDBExplorer'; Gui = $true; Release = @('LICENSE.txt') }
    @{ Name = 'ShellBagsExplorer'; Project = 'ShellBagsExplorer\ShellBagsExplorer2\ShellBagsExplorer2.csproj'; Folder = 'ShellBagsExplorer'; Gui = $true; Release = @('LICENSE.txt') }
    @{ Name = 'TimelineExplorer'; Project = 'timelineexplorer\TimelineExplorer\TimelineExplorer.csproj'; Folder = 'TimelineExplorer'; Gui = $true; Release = @('LICENSE.txt'); Special = 'TlePlugins' }
    @{ Name = 'XWFIM'; Project = 'XWFIM\XWFIM\XWFIM.csproj'; Folder = 'XWFIM'; Gui = $true; Release = @('LICENSE.txt', 'READ ME NOW. No seriously, READ THIS NOW!!!!.txt', 'Urls.txt') }
    @{ Name = 'TimeApp'; Project = 'timeapp\timeapp\timeapp.csproj'; Gui = $true }
    @{ Name = 'Get-ZimmermanTools'; Script = 'get-zimmermantools\Get-ZimmermanTools.ps1' }
)

function Get-MSBuild {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path $vswhere) {
        $path = & $vswhere -latest -prerelease -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
        if ($path) { return $path }
    }
    throw "Visual Studio's MSBuild was not found (it builds the programs with a window)"
}

function Get-SignTool {
    if ($SignTool) {
        if (Test-Path -LiteralPath $SignTool) { return $SignTool }
        throw "signtool.exe not found at '$SignTool'"
    }
    $cmd = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    # the newest Windows SDK's (x64), else the App Certification Kit's, else Visual Studio's ClickOnce one
    $kits = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
    $found = Get-ChildItem $kits -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue | Where-Object { $_.Directory.Name -eq 'x64' } |
        Sort-Object { try { [version]($_.Directory.Parent.Name -replace '[^\d.]', '') } catch { [version]'0.0' } } -Descending | Select-Object -First 1
    if ($found) { return $found.FullName }
    foreach ($candidate in (Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\App Certification Kit\signtool.exe'),
                           (Join-Path ${env:ProgramFiles(x86)} 'Microsoft SDKs\ClickOnce\SignTool\signtool.exe')) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    throw "signtool.exe was not found (Windows SDK); give its path with --signtool"
}

# net472 -> net4, net10.0 / net10.0-windows -> net10, and the old style projects' v4.x -> net4
function Get-Frameworks([string]$project) {
    [xml]$xml = Get-Content -LiteralPath $project -Raw
    $tfms = @($xml.Project.PropertyGroup | ForEach-Object { $_.TargetFrameworks, $_.TargetFramework } | Where-Object { $_ }) -join ';'
    if (-not $tfms) {
        $old = @($xml.Project.PropertyGroup | ForEach-Object { $_.TargetFrameworkVersion } | Where-Object { $_ }) | Select-Object -First 1
        if ($old) { return @([pscustomobject]@{ Tfm = $null; Name = 'net4'; Old = $true }) }
        throw "No target framework in $project"
    }
    foreach ($tfm in $tfms -split ';' | Where-Object { $_ }) {
        $name = if ($tfm -match '^net4') { 'net4' } elseif ($tfm -match '^net(\d+)\.') { "net$($Matches[1])" } else { $tfm }
        [pscustomobject]@{ Tfm = $tfm; Name = $name; Old = $false }
    }
}

function Get-AssemblyName([string]$project) {
    [xml]$xml = Get-Content -LiteralPath $project -Raw
    $name = @($xml.Project.PropertyGroup | ForEach-Object { $_.AssemblyName } | Where-Object { $_ }) | Select-Object -First 1
    if ($name) { return $name }
    return [IO.Path]::GetFileNameWithoutExtension($project)
}

# The current release zips (for the files no repository has), downloaded once
# A note under the program's line (what is downloaded, copied, ...)
function Write-Note([string]$text) {
    Write-Host "   $text" -ForegroundColor DarkGray
}

$releaseCache = $null
function Get-ReleaseFile([string]$program, [string]$framework, [string]$pattern, [string]$folder) {
    if (-not $script:releaseCache) {
        $script:releaseCache = Join-Path $Work '_release'
        New-Item -ItemType Directory -Force $script:releaseCache | Out-Null
        Write-Note "Downloading the download page https://tools.ericzimmermanstools.com (the current release zips' URLs: files no repository has come from them)"
        $page = (Invoke-WebRequest -Uri 'https://tools.ericzimmermanstools.com' -UseBasicParsing).Content
        $regex = [regex] '(?i)\bhttps://[-A-Z0-9+&@#/%?=~_|$!:,.;]*[A-Z0-9+&@#/%=~_|$]\.zip'
        $script:releaseUrls = $regex.Matches($page) | ForEach-Object Value | Sort-Object -Unique
    }

    $net = $framework -ne 'net4'
    $url = $script:releaseUrls | Where-Object { ([uri]$_).Segments[-1] -ieq "$program.zip" -and (($_ -match '/net\d*/') -eq $net) } | Select-Object -First 1
    if (-not $url) {
        # a program that is new for this framework (as Hasher and TimeApp were for .NET 10): the other framework's zip
        $url = $script:releaseUrls | Where-Object { ([uri]$_).Segments[-1] -ieq "$program.zip" } | Select-Object -First 1
    }
    if (-not $url) { Write-Warning "${program}: no release zip to take '$pattern' from"; return @() }

    $zipFile = Join-Path $script:releaseCache (($url -replace '^https://[^/]+/', '') -replace '/', '_')
    if (-not (Test-Path $zipFile)) {
        Write-Note "Downloading the release zip $url (for $pattern)"
        Invoke-WebRequest -Uri $url -OutFile $zipFile -UseBasicParsing
        Write-Note ("   {0:N1} MB saved as {1}" -f ((Get-Item -LiteralPath $zipFile).Length / 1MB), $zipFile)
    }
    else {
        Write-Note "Taking $pattern from the release zip already downloaded: $zipFile"
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($zipFile)
    try {
        $prefix = if ($folder) { "$folder/" } else { '' }
        $like = $prefix + ($pattern -replace '\\', '/')
        $zip.Entries | Where-Object { $_.Name -and $_.FullName -like $like } | ForEach-Object {
            $rel = $_.FullName.Substring($prefix.Length) -replace '/', '\'
            $ms = New-Object IO.MemoryStream
            $s = $_.Open(); $s.CopyTo($ms); $s.Dispose()
            [pscustomobject]@{ Path = $rel; Bytes = $ms.ToArray() }
        }
    }
    finally { $zip.Dispose() }
}

#projects already built with MSBuild (once for all their frameworks)
$msbuildBuilt = @{}

# Builds the program for a framework; returns the folder with its output
# Runs a helper script: its messages (Write-Host: "Copied 7 plugins to ...") indented under the program's line, its
# other output (the build's) dropped
function Invoke-Indented([scriptblock]$block) {
    & $block 6>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.InformationRecord]) { Write-Note $_.MessageData }
    }
}

# A program with a window (.NET): published as one exe (Eric, 2026-10-20), framework dependent, win-x64, its managed
# dependencies inside, native dlls next to it (nothing is extracted to %TEMP%\.net when it runs). TimeApp too (no
# Folder: TimeApp.exe at the top of its zip). Not the command line tools: they stay cross platform (dotnet <Tool>.dll),
# nor net4 builds
function Test-SingleFileGui($p, $fw) {
    return [bool]($p.Gui -and $fw.Name -ne 'net4' -and -not $fw.Old)
}

function Build-Program($p, $fw, [string]$buildDir) {
    $project = Join-Path $Root $p.Project
    if (Test-SingleFileGui $p $fw) {
        # the properties are given here, so the release doesn't depend on what the project file sets
        if (Test-NeedsMSBuild $project $fw) { throw "it needs Visual Studio's MSBuild (licenses.licx), which this script doesn't publish with" }
        $output = & dotnet publish $project -c Release -f $fw.Tfm -r win-x64 --self-contained false -p:PublishSingleFile=true `
            -p:IncludeNativeLibrariesForSelfExtract=false -o $buildDir --nologo -v quiet 2>&1
        if ($LASTEXITCODE -ne 0) { Show-BuildErrors $output; throw "the publish failed" }
        return $buildDir
    }
    if (Test-NeedsMSBuild $project $fw) {
        # no global properties but Configuration: with OutDir or TargetFramework, MSBuild 18 ends with 0x4000001E while
        # it builds a referenced project. One build makes every framework, in the project's bin\Release\<tfm>
        $outputDir = if ($fw.Tfm) { Join-Path (Split-Path $project) "bin\Release\$($fw.Tfm)" } else { Join-Path (Split-Path $project) 'bin\Release' }
        if ($script:msbuildBuilt[$project]) { return $outputDir }
        $msbuild = Get-MSBuild
        $env:MSBUILDDISABLENODEREUSE = '1'
        $run = {
            param([string[]]$a)
            $o = & $msbuild @a 2>&1
            # 0x4000001E: MSBuild worker nodes left from earlier builds; end them (only the workers) and try again
            if ($LASTEXITCODE -eq 0x4000001E) {
                dotnet build-server shutdown 2>&1 | Out-Null
                Get-CimInstance Win32_Process -Filter "Name = 'MSBuild.exe' or Name = 'dotnet.exe'" | Where-Object { $_.CommandLine -match '/nodemode' } |
                    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
                Start-Sleep -Seconds 2
                $o = & $msbuild @a 2>&1
            }
            $o
        }
        # dotnet's MSBuild worker nodes (from the command line tools' builds) also end MSBuild 18 with 0x4000001E
        dotnet build-server shutdown 2>&1 | Out-Null
        # restore, then build (with /restore in the same call, Hasher's build also ended with 0x4000001E)
        $output = & $run @($project, '/t:Restore', '/p:Configuration=Release', '/nr:false', '/v:minimal', '/nologo')
        if ($LASTEXITCODE -eq 0) {
            $output = & $run @($project, '/t:Rebuild', '/p:Configuration=Release', '/nr:false', '/v:minimal', '/nologo')
        }
        if ($LASTEXITCODE -ne 0) { Show-BuildErrors $output; throw "the build failed (exit code $LASTEXITCODE)" }
        $script:msbuildBuilt[$project] = $true
        return $outputDir
    }

    if ($p.SingleFile) {
        # one exe with the dependencies and native libraries (SQLite) inside, as the project's CLAUDE.md releases it
        $output = & dotnet publish $project -c Release -f $fw.Tfm -r win-x64 --self-contained false -p:PublishSingleFile=true -o $buildDir --nologo -v quiet 2>&1
        if ($LASTEXITCODE -ne 0) { Show-BuildErrors $output; throw "the publish failed" }
        return $buildDir
    }

    # a full build: Fody (Costura) only weaves when the compiler runs, so an up-to-date obj left from a build without
    # weaving was copied as is (a 2 MB RegistryExplorer.dll instead of 94 MB with its dependencies)
    $output = & dotnet build $project -c Release -f $fw.Tfm -o $buildDir --no-incremental --nologo -v quiet 2>&1
    if ($LASTEXITCODE -ne 0) { Show-BuildErrors $output; throw "the build failed" }
    return $buildDir
}

# Visual Studio's MSBuild for old-style projects, and for those that embed licenses.licx for any MSBuild (the .NET
# SDK's fails on it: task LC). The others use dotnet build (DevExpress v25+ licenses from the key file at build time)
function Test-NeedsMSBuild([string]$project, $fw) {
    if ($fw.Old) { return $true }
    $text = Get-Content -LiteralPath $project -Raw
    $licx = Get-ChildItem -LiteralPath (Split-Path $project) -Recurse -Filter licenses.licx -File -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' }
    return [bool]$licx -and $text -notmatch "MSBuildRuntimeType"
}

# The dlls Costura embedded in an assembly (lower case names), read from its metadata (any framework)
function Get-CosturaEmbedded([string]$assembly) {
    $names = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $fs = [IO.File]::OpenRead($assembly)
    try {
        $pe = [System.Reflection.PortableExecutable.PEReader]::new($fs)
        if ($pe.HasMetadata) {
            $md = [System.Reflection.Metadata.PEReaderExtensions]::GetMetadataReader($pe)
            foreach ($h in $md.ManifestResources) {
                $n = $md.GetString($md.GetManifestResource($h).Name)
                if ($n -match '^costura\.(.+?\.dll)(\.compressed)?$') { [void]$names.Add($Matches[1]) }
            }
        }
    }
    finally { $fs.Dispose() }
    return ,$names
}

# The files inside a .NET single-file bundle (an exe PublishSingleFile made), from its manifest: the host has the
# manifest's offset in the 8 bytes before a fixed 32-byte signature (0 there: a plain apphost, not a bundle). This is
# what the exe really has (not what a build without single file would make): every bundled file with its path, and
# the bundled deps.json, which lists what the program loads. $null for an exe that is not a bundle.
# Returns {Version; Files = [{Path; Type (Assembly, Native, DepsJson, RuntimeConfigJson, Symbols); Size}]; DepsJson (text)}
function Get-BundleManifest([string]$exe) {
    $signature = [byte[]](0x8b, 0x12, 0x02, 0xb9, 0x6a, 0x61, 0x20, 0x38, 0x72, 0x7b, 0x93, 0x02, 0x14, 0xd7, 0xa0, 0x32,
        0x13, 0xf5, 0xb9, 0xe6, 0xef, 0xae, 0x33, 0x18, 0xee, 0x3b, 0x2d, 0xce, 0x24, 0xb3, 0x6a, 0xae)
    #(one byte per character: a fast ordinal search of the bytes)
    $latin1 = [Text.Encoding]::GetEncoding(28591)
    $fs = [IO.File]::OpenRead($exe)
    try {
        # the signature is in the host, at the start of the file (the bundled files follow it)
        $head = New-Object byte[] ([Math]::Min($fs.Length, 32MB))
        $read = 0
        while ($read -lt $head.Length) { $n = $fs.Read($head, $read, $head.Length - $read); if ($n -le 0) { break }; $read += $n }
        $at = $latin1.GetString($head, 0, $read).IndexOf($latin1.GetString($signature), [StringComparison]::Ordinal)
        if ($at -lt 8) { return $null }
        $headerOffset = [BitConverter]::ToInt64($head, $at - 8)
        if ($headerOffset -le 0 -or $headerOffset -ge $fs.Length) { return $null }

        # the header (major and minor version, file count, bundle id, then since version 2 the deps.json's and
        # runtimeconfig.json's offset and size and the flags), then each file: offset, size, compressed size (version
        # 6+), type, path (a length-prefixed UTF-8 string, as BinaryReader reads it)
        $null = $fs.Seek($headerOffset, [IO.SeekOrigin]::Begin)
        $r = New-Object IO.BinaryReader($fs, [Text.Encoding]::UTF8, $true)
        $major = $r.ReadUInt32(); $null = $r.ReadUInt32()
        $count = $r.ReadInt32(); $null = $r.ReadString()
        $depsOffset = 0; $depsSize = 0
        if ($major -ge 2) {
            $depsOffset = $r.ReadInt64(); $depsSize = $r.ReadInt64()
            $null = $r.ReadInt64(); $null = $r.ReadInt64(); $null = $r.ReadUInt64()
        }
        $types = @('Unknown', 'Assembly', 'Native', 'DepsJson', 'RuntimeConfigJson', 'Symbols')
        $files = for ($i = 0; $i -lt $count; $i++) {
            $null = $r.ReadInt64(); $size = $r.ReadInt64()
            $compressed = if ($major -ge 6) { $r.ReadInt64() } else { 0 }
            $type = $r.ReadByte()
            $path = $r.ReadString() -replace '/', '\'
            [pscustomobject]@{ Path = $path; Type = $(if ($type -lt $types.Count) { $types[$type] } else { "$type" }); Size = $size; Compressed = $compressed }
        }
        $deps = $null
        if ($depsSize -gt 0 -and -not @($files | Where-Object { $_.Type -eq 'DepsJson' -and $_.Compressed -gt 0 })) {
            $null = $fs.Seek($depsOffset, [IO.SeekOrigin]::Begin)
            $deps = [Text.Encoding]::UTF8.GetString($r.ReadBytes([int]$depsSize))
        }
        return [pscustomobject]@{ Version = $major; Files = @($files); DepsJson = $deps }
    }
    finally { $fs.Dispose() }
}

# The files a deps.json says the program loads, relative to its folder as a build lays them out: runtime and native
# assets next to it (by file name), runtimeTargets where their path says (runtimes\<rid>\...), resources in
# <locale>\<file>. For every runtime identifier (the build is for any), so runtimes\unix\... too. -Text: the deps.json's
# text instead of its file (a single file's, from its bundle)
function Get-DepsAssets([string]$depsJson, [string]$Text) {
    $paths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    if (-not $Text) { $Text = Get-Content -LiteralPath $depsJson -Raw }
    $deps = $Text | ConvertFrom-Json
    $libraries = $deps.targets.($deps.runtimeTarget.name)
    foreach ($lib in $libraries.PSObject.Properties) {
        $v = $lib.Value
        foreach ($kind in 'runtime', 'native') {
            if ($v.$kind) { foreach ($a in $v.$kind.PSObject.Properties.Name) { [void]$paths.Add(($a -split '/')[-1]) } }
        }
        if ($v.runtimeTargets) { foreach ($a in $v.runtimeTargets.PSObject.Properties.Name) { [void]$paths.Add(($a -replace '/', '\')) } }
        if ($v.resources) {
            foreach ($r in $v.resources.PSObject.Properties) { [void]$paths.Add("$($r.Value.locale)\$(($r.Name -split '/')[-1])") }
        }
    }
    return @($paths)
}

function Show-BuildErrors($output) {
    $errors = @($output | Select-String -Pattern ' error ' | Select-Object -First 10)
    if (-not $errors) { $errors = @($output | Select-Object -Last 15) }
    $errors | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
}

function Copy-Into([string]$source, [string]$dest) {
    if (Test-Path -LiteralPath $source -PathType Container) {
        New-Item -ItemType Directory -Force $dest | Out-Null
        Copy-Item -Path (Join-Path $source '*') -Destination $dest -Recurse -Force
    }
    else {
        New-Item -ItemType Directory -Force (Split-Path $dest) | Out-Null
        Copy-Item -LiteralPath $source -Destination $dest -Force
    }
}

# A repository's folder or file (<repo>\<path>, under -Root) as committed at HEAD, byte for byte what GitHub has: the
# Windows working copy has CRLF where git stores LF (core.autocrlf), and --sync (EvtxECmd, SQLECmd, RECmd), which
# downloads master from GitHub, compares bytes, so every map looked updated. Uncommitted changes are left out (a
# warning). Returns where the export is
function Export-RepoPath([string]$relative) {
    $parts = $relative -split '[\\/]', 2
    $repo = Join-Path $Root $parts[0]
    $path = $parts[1] -replace '\\', '/'
    if (@(git -C $repo status --porcelain -- $path)) {
        Write-Warning "$relative has uncommitted changes: not staged (the committed files are)"
    }
    $dir = Join-Path $Work "_export\$([guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory -Force $dir | Out-Null
    #(no conversion: core.autocrlf would give the working copy's CRLF again)
    git -C $repo -c core.autocrlf=false -c core.eol=lf archive --format=zip -o "$dir.zip" HEAD -- $path
    if ($LASTEXITCODE -ne 0) { throw "git archive of $relative failed" }
    [IO.Compression.ZipFile]::ExtractToDirectory("$dir.zip", $dir)
    Remove-Item -LiteralPath "$dir.zip"
    $result = Join-Path $dir ($path -replace '/', '\')
    if (-not (Test-Path -LiteralPath $result)) { throw "$relative is not committed in $(Split-Path $repo -Leaf)" }
    Write-Note "$relative as committed ($(Split-Path $repo -Leaf) $(git -C $repo rev-parse --short HEAD))"
    return $result
}

# Microsoft's and DevExpress' files are never signed (--sign checks this too, even with --force): by name, or by the
# company or copyright in their version information
function Test-VendorFile([string]$path) {
    $name = Split-Path $path -Leaf
    if ($name -match '^(Microsoft|System|DevExpress|mscorlib|netstandard|WindowsBase|Presentation(Core|Framework))\.') { return $true }
    if ($name -notmatch '\.(exe|dll)$') { return $false }
    try { $vi = [Diagnostics.FileVersionInfo]::GetVersionInfo($path) } catch { return $false }
    #(not "Microsoft" anywhere: a project without a Company has its assembly name there, such as TLEFileMicrosoft)
    return "$($vi.CompanyName)" -match '^\s*(Microsoft|Developer Express|DevExpress)\b' -or
        "$($vi.LegalCopyright)" -match 'Microsoft Corporation|Developer Express'
}

# Is a NuGet package Eric's? Its .nuspec in the NuGet cache names him (authors, owners or copyright), or its project or
# repository is on github.com/EricZimmerman (older ShellBagsParser and ERZHelpers packages have the package id as
# authors). Without the .nuspec (a cache cleared since the build): the known package ids, with a warning
$nugetRoot = if ($env:NUGET_PACKAGES) { $env:NUGET_PACKAGES } else { Join-Path $env:USERPROFILE '.nuget\packages' }
$ericPackageCache = @{}
$ericPackagesFallback = @('Amcache', 'AppCompatCache', 'Boot', 'ERZHelpers', 'ExtensionBlocks', 'EZ.HexBox', 'GuidMapping',
    'HexEditorControl', 'I30', 'JumpList', 'Lnk', 'MFT', 'MFT.LogFile', 'OleCf', 'Prefetch', 'Registry', 'RegistryPluginBase',
    'SDB', 'Secure', 'ShellBagsParser', 'Usn')
function Test-EricPackage([string]$id, [string]$path) {
    $key = "$id/$path"
    if ($script:ericPackageCache.ContainsKey($key)) { return $script:ericPackageCache[$key] }
    $nuspec = if ($path) { Get-ChildItem -LiteralPath (Join-Path $nugetRoot $path) -Filter *.nuspec -File -ErrorAction SilentlyContinue | Select-Object -First 1 }
    if ($nuspec) {
        $text = [IO.File]::ReadAllText($nuspec.FullName)
        $mine = $text -match '<(authors|owners|copyright)>[^<]*Zimmerman' -or $text -match '(?i)github\.com/EricZimmerman/'
    }
    else {
        $mine = $ericPackagesFallback -contains $id
        Write-Warning "No .nuspec for $id in $nugetRoot ($path): $(if ($mine) { 'Eric''s by the known list' } else { 'taken as third-party (not signed)' })"
    }
    $script:ericPackageCache[$key] = $mine
    return $mine
}

# The files of a deps.json (relative to its folder, as Get-DepsAssets lays them out) that are Eric's: from a project
# (his repositories) or from a package of his
function Get-EricDepsAssets([string]$depsJson) {
    $paths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $deps = Get-Content -LiteralPath $depsJson -Raw | ConvertFrom-Json
    $libraries = $deps.targets.($deps.runtimeTarget.name)
    foreach ($lib in $libraries.PSObject.Properties) {
        $info = $deps.libraries.($lib.Name)
        $mine = if ($info.type -eq 'project') { $true }
            elseif ($info.type -eq 'package') { Test-EricPackage (($lib.Name -split '/')[0]) $info.path }
            else { $false }
        if (-not $mine) { continue }
        $v = $lib.Value
        foreach ($kind in 'runtime', 'native') {
            if ($v.$kind) { foreach ($a in $v.$kind.PSObject.Properties.Name) { [void]$paths.Add(($a -split '/')[-1]) } }
        }
        if ($v.runtimeTargets) { foreach ($a in $v.runtimeTargets.PSObject.Properties.Name) { [void]$paths.Add(($a -replace '/', '\')) } }
        if ($v.resources) {
            foreach ($r in $v.resources.PSObject.Properties) { [void]$paths.Add("$($r.Value.locale)\$(($r.Name -split '/')[-1])") }
        }
    }
    return ,$paths
}

# What --sign signs in a staged program (<Out>\<Program>\<framework>), decided here, where the build and the NuGet cache
# are (the signing machine has neither): every .exe and .ps1, and the dlls that are Eric's (the program's own, the
# plugins, his projects and packages). Third-party, Microsoft and DevExpress files are left out. Returns
# {Path (relative to the stage); Sign; Why}
function Get-SignEntries([string]$stage, [string]$target, [string]$asm, [string]$depsJson, $pluginNames) {
    $own = if ($depsJson -and (Test-Path -LiteralPath $depsJson)) { Get-EricDepsAssets $depsJson } else { $null }
    foreach ($f in Get-ChildItem -LiteralPath $stage -Recurse -File | Where-Object { $_.Extension -in '.exe', '.dll', '.ps1' }) {
        $rel = $f.FullName.Substring($stage.TrimEnd('\').Length + 1)
        $relTarget = if ($f.FullName.StartsWith("$($target.TrimEnd('\'))\", [StringComparison]::OrdinalIgnoreCase)) { $f.FullName.Substring($target.TrimEnd('\').Length + 1) } else { $rel }
        $why = if (Test-VendorFile $f.FullName) { $null }
            elseif ($f.Extension -in '.exe', '.ps1') { $f.Extension }
            elseif ($relTarget -ieq "$asm.dll") { 'the program' }
            elseif ($relTarget -like 'Plugins\*' -and $pluginNames -contains $f.Name) { 'plugin' }
            elseif ($own -and $own.Contains($relTarget)) { 'Eric''s project or package' }
            else { $null }
        [pscustomobject]@{ Path = $rel; Sign = [bool]$why; Why = $why }
    }
}

function Invoke-Stage {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $Out = [IO.Path]::GetFullPath($Out)
    New-Item -ItemType Directory -Force $Out | Out-Null
    $buildRoot = Join-Path $Work '_build'
    if (Test-Path -LiteralPath (Join-Path $Work '_export')) { Remove-Item -LiteralPath (Join-Path $Work '_export') -Recurse -Force }
    # left in -Out by earlier versions of this script
    foreach ($old in '_build', '_release') {
        $oldDir = Join-Path $Out $old
        if (Test-Path -LiteralPath $oldDir) {
            Write-Host "Deleting $oldDir (the build output and release zips are in $Work now)" -ForegroundColor Yellow
            Remove-Item -LiteralPath $oldDir -Recurse -Force
        }
    }
    $results = New-Object System.Collections.Generic.List[object]

    $selected = if ($Only) { $programs | Where-Object { $Only -contains $_.Name } } else { $programs }
    if ($Only) {
        $unknown = $Only | Where-Object { $programs.Name -notcontains $_ }
        if ($unknown) { throw "Unknown program(s): $($unknown -join ', '). Known: $($programs.Name -join ', ')" }
    }

    # -Root defaults to the folder above this script: a copy run from elsewhere (a share) finds no repositories
    $missingRoot = @($selected | Where-Object { -not (Test-Path -LiteralPath (Join-Path $Root $(if ($_.Script) { $_.Script } else { $_.Project }))) })
    if ($missingRoot.Count -eq @($selected).Count) {
        throw "None of the repositories are in '$Root' (the folder above this script, unless --root is given). Run the script from its repository (D:\Code\get-zimmermantools) or give --root <the folder with the repositories>, such as --root D:\Code"
    }

    $registryPlugins = $null
    $tlePlugins = $null
    # the files --sign signs (written to <Out>\_sign.txt), and the staged programs this run replaced (Program\fw)
    $signList = New-Object System.Collections.Generic.List[string]
    $stagedKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

    foreach ($p in $selected) {
        # a script (Get-ZimmermanTools.ps1): copied as it is into <Out>\<Name>\script, signed there by --sign (the
        # repository's copy stays unsigned) and zipped at the top
        if ($p.Script) {
            Write-Host '== ' -ForegroundColor Red -NoNewline
            Write-Host "$($p.Name) script" -ForegroundColor Green
            $src = Join-Path $Root $p.Script
            $dest = Join-Path $Out "$($p.Name)\script"
            [void]$stagedKeys.Add("$($p.Name)\script")
            if (-not (Test-Path -LiteralPath $src)) {
                Write-Host "   $($p.Name): $src not found" -ForegroundColor Red
                $results.Add([pscustomobject]@{ Program = $p.Name; Framework = 'script'; Result = "FAILED: $($p.Script) not found" })
                continue
            }
            if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Recurse -Force }
            Copy-Into $src (Join-Path $dest (Split-Path $p.Script -Leaf))
            Write-Note "Copied $src"
            $signList.Add("$($p.Name)\script\$(Split-Path $p.Script -Leaf)")
            $results.Add([pscustomobject]@{ Program = $p.Name; Framework = 'script'; Result = 'staged' })
            continue
        }

        $project = Join-Path $Root $p.Project
        if (-not (Test-Path $project)) {
            Write-Host '== ' -ForegroundColor Red -NoNewline
            Write-Host $p.Name -ForegroundColor Green
            Write-Host "   $($p.Name): $project not found" -ForegroundColor Red
            $results.Add([pscustomobject]@{ Program = $p.Name; Framework = '-'; Result = "no project ($($p.Project))" })
            continue
        }

        $asm = Get-AssemblyName $project
        foreach ($fw in Get-Frameworks $project) {
            Write-Host '== ' -ForegroundColor Red -NoNewline
            Write-Host "$($p.Name) $($fw.Name)" -ForegroundColor Green
            $buildDir = Join-Path $buildRoot "$($p.Name)\$($fw.Name)"
            $stage = Join-Path $Out "$($p.Name)\$($fw.Name)"
            $target = if ($p.Folder) { Join-Path $stage $p.Folder } else { $stage }
            [void]$stagedKeys.Add("$($p.Name)\$($fw.Name)")

            try {
                if (Test-Path $buildDir) { Remove-Item -LiteralPath $buildDir -Recurse -Force }
                if (Test-Path $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
                $buildDir = Build-Program $p $fw $buildDir
                New-Item -ItemType Directory -Force $target | Out-Null

                $singleGui = Test-SingleFileGui $p $fw
                $costura = (Get-Content -LiteralPath $project -Raw) -match 'Costura\.Fody'
                $depsJson = Join-Path $buildDir "$asm.deps.json"
                if ($singleGui) {
                    # a program with a window (Eric, 2026-10-20): the publish output as it is, the exe (its managed
                    # dependencies inside) and what the publish leaves next to it (native dlls such as sni.dll,
                    # {0}.dll.config, content files), in its own Folder; no pdbs (the release zips never had any)
                    if ($costura) { throw "the project still has Costura.Fody (a program with a window is a single file without it)" }
                    foreach ($f in @('{0}.exe') + @($p.Files) | Where-Object { $_ }) {
                        if (-not (Test-Path -LiteralPath (Join-Path $buildDir ($f -f $asm)))) { throw "'$($f -f $asm)' is not in the publish output" }
                    }
                    $published = @(Get-ChildItem -LiteralPath $buildDir -Recurse -File | Where-Object Extension -ne '.pdb')
                    foreach ($f in $published) {
                        $dest = Join-Path $target $f.FullName.Substring($buildDir.TrimEnd('\').Length + 1)
                        New-Item -ItemType Directory -Force (Split-Path $dest) | Out-Null
                        Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
                    }
                    $depsJson = $null
                }
                else {
                    #(@(): a one-item array from an if is unrolled to a string, and += then appends to the string)
                    $files = @(if ($fw.Name -eq 'net4' -or $p.SingleFile) { '{0}.exe' } else { '{0}.exe', '{0}.dll', '{0}.runtimeconfig.json' })
                    $files += @($p.Files)
                    foreach ($f in $files | Where-Object { $_ }) {
                        $name = $f -f $asm
                        $src = Join-Path $buildDir $name
                        if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination $target -Force }
                        elseif ($name -notlike '*.runtimeconfig.json' -or $fw.Name -ne 'net4') { throw "'$name' is not in the build output" }
                    }
                }

                foreach ($k in @($p.Copy.Keys)) {
                    if ($k) { Copy-Into (Export-RepoPath $k) (Join-Path $target $p.Copy[$k]) }
                }

                foreach ($d in @($p.EmptyDirs)) {
                    if ($d) { New-Item -ItemType Directory -Force (Join-Path $target $d) | Out-Null }
                }

                foreach ($pattern in @($p.Release)) {
                    if (-not $pattern) { continue }
                    $got = @(Get-ReleaseFile $p.Name $fw.Name $pattern $p.Folder)
                    if (-not $got) { Write-Warning "$($p.Name) $($fw.Name): '$pattern' not found in the release zip" }
                    foreach ($g in $got) {
                        $dest = Join-Path $target $g.Path
                        New-Item -ItemType Directory -Force (Split-Path $dest) | Out-Null
                        [IO.File]::WriteAllBytes($dest, $g.Bytes)
                    }
                }

                switch ($p.Special) {
                    'RegistryPlugins' {
                        if (-not $registryPlugins) {
                            Invoke-Indented { & (Join-Path $Root 'RegistryPlugins\CopyPlugins.ps1') -Out (Join-Path $buildRoot 'RegistryPlugins') -Build }
                            $registryPlugins = Join-Path $buildRoot 'RegistryPlugins'
                        }
                        Copy-Into $registryPlugins (Join-Path $target 'Plugins')
                    }
                    'TlePlugins' {
                        if (-not $tlePlugins) {
                            $tlePlugins = Join-Path $buildRoot 'TLEFilePlugins'
                            Invoke-Indented { & (Join-Path $Root 'tlefileplugins\CopyPlugins.ps1') -Out (Join-Path $tlePlugins 'Plugins') -Build }
                            # the layouts as committed (CopyPlugins -Layouts copies the working copy's)
                            Copy-Into (Export-RepoPath 'tlefileplugins\Layouts\Default') (Join-Path $tlePlugins 'Layouts\Default')
                            Copy-Into (Export-RepoPath 'tlefileplugins\Layouts\!!README.txt') (Join-Path $tlePlugins 'Layouts\!!README.txt')
                        }
                        Copy-Into (Join-Path $tlePlugins 'Plugins') (Join-Path $target 'Plugins')
                        Copy-Into (Join-Path $tlePlugins 'Layouts') (Join-Path $target 'Layouts')
                    }
                }

                # every dll the build made must be in the program: staged next to it, or embedded by Costura or in the
                # single file's bundle (a build without either can't run: the dependencies would be missing)
                $embedded = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                $main = Join-Path $target ($(if ($fw.Name -eq 'net4' -or $p.SingleFile -or $singleGui) { '{0}.exe' } else { '{0}.dll' }) -f $asm)
                if (Test-Path -LiteralPath $main) { $embedded.UnionWith((Get-CosturaEmbedded $main)) }
                if ($costura -and $embedded.Count -eq 0) {
                    throw "$(Split-Path $main -Leaf) has no embedded dependencies (Costura did not run)"
                }

                if ($singleGui) {
                    # the exe is a single file: its bundle (read from its manifest, so this is what the exe really has)
                    # has the managed dependencies and the deps.json that lists everything the program loads
                    $bundle = Get-BundleManifest $main
                    if (-not $bundle) { throw "$(Split-Path $main -Leaf) is not a single file (no bundle in it)" }
                    if (-not $bundle.DepsJson) { throw "$(Split-Path $main -Leaf) has no deps.json in its bundle" }
                    foreach ($b in $bundle.Files) { [void]$embedded.Add((Split-Path $b.Path -Leaf)) }
                    $inBundle = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                    foreach ($b in $bundle.Files) { [void]$inBundle.Add($b.Path) }

                    # every file the deps.json lists is in the bundle or next to the exe (it doesn't start without them)
                    $missingDeps = @(Get-DepsAssets -Text $bundle.DepsJson | Where-Object {
                            -not $inBundle.Contains($_) -and -not (Test-Path -LiteralPath (Join-Path $target $_)) })
                    if ($missingDeps) {
                        throw "$($missingDeps.Count) file(s) its deps.json lists are neither in the exe nor next to it: $(($missingDeps | Select-Object -First 4) -join ', ')"
                    }

                    # every file the publish made (native dlls, the config, content) is staged; those next to the exe as
                    # published. In subfolders the table's steps may have replaced them: Timeline Explorer's project
                    # puts its plugins and the working copy's layouts next to the publish, Special stages the plugins
                    # CopyPlugins built and the layouts as committed
                    $notStaged = @($published | Where-Object {
                            $rel = $_.FullName.Substring($buildDir.TrimEnd('\').Length + 1)
                            $dest = Join-Path $target $rel
                            -not (Test-Path -LiteralPath $dest) -or ($rel -notmatch '\\' -and (Get-Item -LiteralPath $dest).Length -ne $_.Length)
                        } | ForEach-Object Name)
                    if ($notStaged) {
                        throw "$($notStaged.Count) file(s) of the publish are not staged as published: $(($notStaged | Select-Object -First 4) -join ', ')"
                    }

                    # no managed dll loose in the program (it belongs in the exe; a copy next to it could be loaded
                    # instead): only native ones. The plugins (Plugins\) are loaded from their files
                    $loose = @(Get-ChildItem -LiteralPath $target -Recurse -File -Filter *.dll | Where-Object {
                            $_.FullName.Substring($target.TrimEnd('\').Length + 1) -notlike 'Plugins\*' } | Where-Object {
                            $fs = [IO.File]::OpenRead($_.FullName)
                            try { ([System.Reflection.PortableExecutable.PEReader]::new($fs)).HasMetadata } catch { $false } finally { $fs.Dispose() }
                        })
                    if ($loose) {
                        throw "$($loose.Count) managed dll(s) next to the single file (they belong inside it): $(($loose | ForEach-Object { $_.FullName.Substring($target.TrimEnd('\').Length + 1) } | Select-Object -First 4) -join ', ')"
                    }
                    $natives = @(Get-ChildItem -LiteralPath $target -File -Filter *.dll | ForEach-Object Name)
                    Write-Note ("Single file: {0} ({1:N1} MB, {2} files inside); next to it: {3}" -f (Split-Path $main -Leaf),
                        ((Get-Item -LiteralPath $main).Length / 1MB), $bundle.Files.Count, $(if ($natives) { $natives -join ', ' } else { 'no native dll' }))
                }

                #(staged anywhere in the program: Timeline Explorer's plugins are in Plugins)
                $stagedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                Get-ChildItem -LiteralPath $stage -Recurse -File -Filter *.dll | ForEach-Object { [void]$stagedNames.Add($_.Name) }
                $missing = @(Get-ChildItem $buildDir -Filter *.dll -File | Where-Object {
                    -not $stagedNames.Contains($_.Name) -and -not $embedded.Contains($_.Name) })
                if ($missing) {
                    throw "$($missing.Count) dll(s) of the build are neither embedded nor staged: $(($missing.Name | Select-Object -First 4) -join ', ')"
                }
                # what --sign signs (decided here: the signing machine has no build output or NuGet cache)
                $pluginNames = @(switch ($p.Special) {
                        'RegistryPlugins' { Get-ChildItem -LiteralPath $registryPlugins -Filter *.dll -File | ForEach-Object Name }
                        'TlePlugins' { Get-ChildItem -LiteralPath (Join-Path $tlePlugins 'Plugins') -Filter *.dll -File | ForEach-Object Name }
                    })
                $entries = @(Get-SignEntries $stage $target $asm $depsJson $pluginNames)
                $toSign = @($entries | Where-Object Sign)
                foreach ($e in $toSign) { $signList.Add("$($p.Name)\$($fw.Name)\$($e.Path)") }
                Write-Note ("To sign: {0} of {1} .exe/.dll/.ps1 (Eric's); {2} third-party, Microsoft or DevExpress left unsigned" -f
                    $toSign.Count, $entries.Count, ($entries.Count - $toSign.Count))

                $note = if ($p.SingleFile -or $singleGui) { 'staged (single file)' } else { 'staged' }
                $results.Add([pscustomobject]@{ Program = $p.Name; Framework = $fw.Name; Result = $note })
            }
            catch {
                # nothing left of a failed program: --zip would replace its download zip with what was staged so far
                if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
                Write-Host "   $($p.Name) $($fw.Name): $($_.Exception.Message)" -ForegroundColor Red
                $results.Add([pscustomobject]@{ Program = $p.Name; Framework = $fw.Name; Result = "FAILED: $($_.Exception.Message)" })
            }
        }
    }

    $table = $results | Format-Table -AutoSize | Out-String -Width 220
    Write-Host $table
    $bad = @($results | Where-Object { $_.Result -like 'FAILED*' -or $_.Result -like 'no project*' })
    if ($bad) { $script:failed = $true }

    # _sign.txt: the files --sign signs, one path per line (relative to -Out). The lines of programs not staged this run
    # (an earlier run with other --tool names) are kept while their files are there
    $manifest = Join-Path $Out '_sign.txt'
    $kept = @()
    if (Test-Path -LiteralPath $manifest) {
        $kept = @(Get-Content -LiteralPath $manifest | Where-Object { $_ -and $_ -notmatch '^\s*#' } | Where-Object {
                $parts = $_ -split '\\'
                $parts.Count -gt 2 -and -not $stagedKeys.Contains("$($parts[0])\$($parts[1])") -and (Test-Path -LiteralPath (Join-Path $Out $_))
            })
    }
    $all = @($kept) + @($signList) | Sort-Object -Unique
    @(
        "# The files Build-ZimmermanTools.ps1 --sign signs (paths relative to this folder), written by staging on $env:COMPUTERNAME"
        "# $(Get-Date -Format 'yyyy-MM-dd HH:mm'): every .exe and .ps1 in the programs, and the dlls that are Eric's (each program's"
        "# own, the plugins, his projects and his NuGet packages). Third-party, Microsoft and DevExpress files are not listed."
        "# Not in any zip. Restage (--out) to change it."
        $all
    ) | Set-Content -LiteralPath $manifest -Encoding UTF8
    Write-Note "$manifest`: $($all.Count) file(s) to sign"

    # everything another machine needs to sign and zip: this script, and what was staged (with the commands)
    $self =Join-Path $Out (Split-Path $PSCommandPath -Leaf)
    if ($self -ne $PSCommandPath) { Copy-Item -LiteralPath $PSCommandPath -Destination $self -Force }
    $stagedNow = Get-ChildItem -LiteralPath $Out -Directory | Where-Object { $_.Name -notlike '_*' } | ForEach-Object {
        "  $($_.Name): $((Get-ChildItem -LiteralPath $_.FullName -Directory).Name -join ', ')" }
    @(
        "Staged $(Get-Date -Format 'yyyy-MM-dd HH:mm') on $env:COMPUTERNAME from $([IO.Path]::GetFullPath($Root))"
        $bad | ForEach-Object { "FAILED $($_.Program) $($_.Framework): $($_.Result)" }
        ''
        'This run:'
        $table.Trim()
        ''
        'In this folder:'
        $stagedNow
        ''
        'To sign and zip (on any machine; no repositories needed), from this folder:'
        "  .\$(Split-Path $PSCommandPath -Leaf) --sign --zip <download folder>"
        "  .\$(Split-Path $PSCommandPath -Leaf) --sign                     (sign only; --zip <folder> later)"
        '  (--sign signs only the files _sign.txt lists: Eric''s, not third-party, Microsoft or DevExpress files)'
    ) | Set-Content -LiteralPath (Join-Path $Out '_staged.txt') -Encoding UTF8
    Write-Host "To sign and zip on another machine: $self --sign --zip <download folder>" -ForegroundColor Yellow
}

# The files whose Authenticode signature is not valid (each file is read in full and its certificate chain checked:
# about 10 s for a full release; several at a time in PowerShell 7)
function Get-UnsignedFiles($files, [string]$activity) {
    $files = @($files)
    if ($files.Count -eq 0) { return @() }
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $paths = $files.FullName | ForEach-Object -ThrottleLimit 8 -Parallel {
            if ((Get-AuthenticodeSignature -LiteralPath $_).Status -ne 'Valid') { $_ }
        }
        return @($files | Where-Object { $paths -contains $_.FullName })
    }
    $i = 0
    $result = foreach ($f in $files) {
        Write-Progress -Activity $activity -Status $f.Name -PercentComplete (100 * $i++ / $files.Count)
        if ((Get-AuthenticodeSignature -LiteralPath $f.FullName).Status -ne 'Valid') { $f }
    }
    Write-Progress -Activity $activity -Completed
    return @($result)
}

function Invoke-Sign([string]$dir) {
    $dir = [IO.Path]::GetFullPath($dir).TrimEnd('\')
    # only the files staging listed in _sign.txt (Eric's: not third-party, Microsoft or DevExpress files); a folder
    # staged before it existed is not signed at all rather than everything in it
    $manifest = Join-Path $dir '_sign.txt'
    if (-not (Test-Path -LiteralPath $manifest)) {
        Write-Host "Not signing: $dir has no _sign.txt (the list of Eric's files to sign, written by staging). It was staged by an older version of this script, or is not a staged folder (give the folder --out staged to, not a program in it): restage it with --out." -ForegroundColor Red
        $script:failed = $true
        $script:signFailed = $true
        return
    }
    Write-Host "Reading the files to sign from $manifest"
    $listed = @(Get-Content -LiteralPath $manifest | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') } | Sort-Object -Unique)
    $missingFiles = @($listed | Where-Object { -not (Test-Path -LiteralPath (Join-Path $dir $_) -PathType Leaf) })
    if ($missingFiles) {
        Write-Host ("{0:N0} file(s) in _sign.txt are not in {1} (not signed):" -f $missingFiles.Count, $dir) -ForegroundColor Red
        $missingFiles | Select-Object -First 20 | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        if ($missingFiles.Count -gt 20) { Write-Host "  ... and $($missingFiles.Count - 20) more" -ForegroundColor Red }
        $script:failed = $true
        $script:signFailed = $true
    }
    $files = @($listed | Where-Object { $missingFiles -notcontains $_ } | ForEach-Object { Get-Item -LiteralPath (Join-Path $dir $_) })
    # never Microsoft's or DevExpress' (not even with --force), whatever the list says
    $vendor = @($files | Where-Object { Test-VendorFile $_.FullName })
    if ($vendor) {
        Write-Warning "$($vendor.Count) file(s) in _sign.txt are Microsoft's or DevExpress' (not signed): $(($vendor.Name | Select-Object -First 5) -join ', ')"
        $files = @($files | Where-Object { $vendor.FullName -notcontains $_.FullName })
    }

    # the other .exe/.dll/.ps1 files in the programs (third-party, Microsoft, DevExpress): left unsigned by design
    $listedSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($l in $listed) { [void]$listedSet.Add([IO.Path]::GetFullPath((Join-Path $dir $l))) }
    $others = @(Get-ChildItem -LiteralPath $dir -Recurse -File | Where-Object { $_.Extension -in '.exe', '.dll', '.ps1' } |
        Where-Object { $_.DirectoryName.TrimEnd('\') -ine $dir -and -not $listedSet.Contains($_.FullName) })
    $total = $files.Count
    Write-Note ("{0:N0} file(s) to sign, {1:N0} MB; {2:N0} other .exe/.dll/.ps1 file(s) left unsigned by design (third-party, Microsoft, DevExpress)" -f
        $total, (($files | Measure-Object Length -Sum).Sum / 1MB), $others.Count)
    # a staged program with nothing listed was staged before _sign.txt (or by a run that failed for it)
    $unlisted = @(Get-ChildItem -LiteralPath $dir -Directory | Where-Object { $_.Name -notlike '_*' } | ForEach-Object {
            foreach ($fwDir in Get-ChildItem -LiteralPath $_.FullName -Directory) {
                $prefix = "$($_.Name)\$($fwDir.Name)\"
                if (-not ($listed | Where-Object { $_.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) })) { "$($_.Name)\$($fwDir.Name)" }
            } })
    if ($unlisted) { Write-Warning "Nothing in _sign.txt for $($unlisted -join ', '): not signed (restage them with --out)" }

    if (-not $Force -and $total -gt 0) {
        Write-Note "Checking their signatures (files already validly signed are skipped; --force signs them again)"
        $files = Get-UnsignedFiles $files 'Checking signatures'
        Write-Note ("{0:N0} to sign, {1:N0} already signed" -f $files.Count, ($total - $files.Count))
    }
    $files = @($files)
    if ($files.Count -eq 0) { Write-Host "Nothing to sign in $dir"; return }
    $signtool = Get-SignTool

    $base = @('sign', '/fd', 'sha256', '/tr', $TimestampUrl, '/td', 'sha256')
    $plain = $null
    if ($CertFile) {
        $secure = Read-Host -AsSecureString "Password for $CertFile"
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
        $base += @('/f', $CertFile, '/p', $plain)
    }
    else {
        $base += '/a'
    }

    Write-Host "Signing $($files.Count) file(s) in $dir with $signtool"

    # the files of each call (a token asks for its password once per call): as many as the command line takes (32,767 characters for an exe; a .cmd or .bat wrapper
    # goes through cmd.exe: 8,191)
    $max = if ($signtool -match '\.(cmd|bat)$') { 7000 } else { 28000 }
    $batches = New-Object System.Collections.Generic.List[object]
    $current = New-Object System.Collections.Generic.List[string]
    $length = 0
    foreach ($f in $files) {
        if ($current.Count -gt 0 -and $length + $f.FullName.Length + 3 -gt $max) {
            $batches.Add($current); $current = New-Object System.Collections.Generic.List[string]; $length = 0
        }
        $current.Add($f.FullName)
        $length += $f.FullName.Length + 3
    }
    if ($current.Count -gt 0) { $batches.Add($current) }

    for ($b = 0; $b -lt $batches.Count; $b++) {
        Write-Note ("Signing batch {0} of {1} ({2:N0} files){3}" -f ($b + 1), $batches.Count, $batches[$b].Count,
            $(if (-not $CertFile) { ': a token asks for its password now' } else { '' }))
        & $signtool @base @($batches[$b])
        if ($LASTEXITCODE -ne 0) { $script:signFailed = $true }
    }
    $plain = $null

    Write-Host "Verifying the signatures"
    $bad = @(Get-UnsignedFiles $files 'Verifying signatures')
    if ($bad) {
        Write-Host ("Not signed (or not valid): {0:N0} of {1:N0} file(s)" -f $bad.Count, $files.Count) -ForegroundColor Red
        $bad | Select-Object -First 20 | ForEach-Object { Write-Host "  $($_.FullName)" -ForegroundColor Red }
        if ($bad.Count -gt 20) { Write-Host "  ... and $($bad.Count - 20) more" -ForegroundColor Red }
        $script:failed = $true
        $script:signFailed = $true
    }
    else {
        Write-Host "All $($files.Count) file(s) signed" -ForegroundColor Green
    }
}

function Invoke-Zip([string]$staged, [string]$dest) {
    $staged = [IO.Path]::GetFullPath($staged)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $count = 0
    foreach ($programDir in Get-ChildItem -LiteralPath $staged -Directory | Where-Object { $_.Name -notlike '_*' }) {
        $p = $programs | Where-Object { $_.Name -eq $programDir.Name } | Select-Object -First 1
        $zipName = if ($p -and $p.ZipName) { $p.ZipName } else { $programDir.Name }
        foreach ($fwDir in Get-ChildItem -LiteralPath $programDir.FullName -Directory) {
            # .NET Framework builds and the script at the top, the others in net\
            $folder = if ($fwDir.Name -in @('net4', 'script')) { $dest } else { Join-Path $dest 'net' }
            New-Item -ItemType Directory -Force $folder | Out-Null
            $zipFile = Join-Path $folder "$zipName.zip"

            # made locally, then copied (the destination can be a share)
            $tmp = Join-Path ([IO.Path]::GetTempPath()) "$zipName-$($fwDir.Name)-$([guid]::NewGuid().ToString('N')).zip"
            [IO.Compression.ZipFile]::CreateFromDirectory($fwDir.FullName, $tmp, [IO.Compression.CompressionLevel]::Optimal, $false)

            # copied next to the zip under another name, then renamed over it: a download while it is copied gets the
            # whole old zip, never part of the new one (a rename in the same folder replaces the file in one step)
            $partial = "$zipFile.partial-$([guid]::NewGuid().ToString('N'))"
            try {
                Copy-Item -LiteralPath $tmp -Destination $partial -Force
                # (Windows PowerShell 5.1's .NET Framework has no Move with overwrite: Replace does the same in one step)
                if ($PSVersionTable.PSVersion.Major -ge 7) { [IO.File]::Move($partial, $zipFile, $true) }
                #(no backup: [NullString]::Value, since $null arrives as "")
                elseif (Test-Path -LiteralPath $zipFile) { [IO.File]::Replace($partial, $zipFile, [NullString]::Value) }
                else { [IO.File]::Move($partial, $zipFile) }
            }
            finally {
                if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Force }
                Remove-Item -LiteralPath $tmp -Force
            }
            Write-Host "  $zipFile"
            $count++
        }

        # a zip left from a framework the program isn't built for any more (Hasher, AmcacheParser and XWFIM have no
        # .NET Framework build): reported, not deleted (the destination is often the download share)
        $frameworks = @(Get-ChildItem -LiteralPath $programDir.FullName -Directory | ForEach-Object Name)
        $hasTop = @($frameworks | Where-Object { $_ -in @('net4', 'script') }).Count -gt 0
        $hasNet = @($frameworks | Where-Object { $_ -notin @('net4', 'script') }).Count -gt 0
        foreach ($stale in @(
                $(if (-not $hasTop) { Join-Path $dest "$zipName.zip" }),
                $(if (-not $hasNet) { Join-Path (Join-Path $dest 'net') "$zipName.zip" })) | Where-Object { $_ }) {
            if (Test-Path -LiteralPath $stale) {
                Write-Warning "$stale is left from a build $($programDir.Name) doesn't have any more: delete it (and its link on the download page)"
            }
        }
    }
    Write-Host "$count zip(s) in $dest"
}

$script:failed = $false
$script:signFailed = $false
if ($stagedFrom) {
    $lines = @(Get-Content -LiteralPath $stagedFrom)
    Write-Host "Staged folder $Staged ($($lines[0]))"
    $lines | Where-Object { $_ -like 'FAILED *' } | ForEach-Object { Write-Host "   $_ (not staged: not in the zips)" -ForegroundColor Red }
}
if ($Out) { Invoke-Stage }
if ($Sign) { Invoke-Sign $Sign }
# zips only with every file signed: a signing failure leaves the download folder as it was
if ($Zip -and $Sign -and $script:signFailed) { Write-Host "Not zipping: signing failed" -ForegroundColor Red }
elseif ($Zip) { Invoke-Zip $(if ($Out) { $Out } else { $Staged }) $Zip }
if ($script:failed) { exit 1 }
