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
      - Maps, batch files, bookmarks, the manual and the layouts are taken as committed (HEAD), byte for byte what
        GitHub has (and what --sync downloads), not from the working copy (CRLF on Windows); uncommitted changes in
        them are left out with a warning
      - Files that are not in any repository (the GUIs' LICENSE.txt, iisGeolocate's GeoLite2-City.mmdb, Registry
        Explorer's default Settings, XWFIM's text files) are taken from the current release zips, downloaded once
        (from the page Get-ZimmermanTools.ps1 reads) into %TEMP%\Build-ZimmermanTools\_release

    The build output is in %TEMP%\Build-ZimmermanTools\_build: -Out (which can be a share) only gets the staged
    programs, a copy of this script and _staged.txt. The copy is replaced by every staging run, so the one in -Out is
    always the current one.

    Command line tools are built with dotnet, programs with a window with Visual Studio's MSBuild (DevExpress licensing
    and the old-style projects need it).

    Get-ZimmermanTools.ps1 is staged too (<Out>\Get-ZimmermanTools\script), so it is signed and zipped with the rest;
    -Only Get-ZimmermanTools stages just the script.

    Signing (-Sign <dir>, or --sign): signs every .exe, .dll and .ps1 under the folder with signtool.exe
    (sign /fd sha256 /tr <timestamp> /td sha256), many files per call, so a token asks for its password once per call:
    with the certificate picked automatically (/a, the default), or from a .pfx (-CertFile; its password is asked once).
    Files that already have a valid signature (DevExpress, Microsoft) are left alone unless -Force is used.

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
    A folder to sign, recursively (.exe, .dll and .ps1). Also --sign. Without a value after -Out, the -Out folder.

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
    Sign files that already have a valid signature too.

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
# others: {0}.exe, {0}.dll, {0}.runtimeconfig.json). Copy: repository folders/files -> paths in the staged program.
# Release: files only in the release zips. Special: steps for plugins and layouts. SingleFile: published as one exe
# (framework dependent, win-x64, native libraries inside: the tools with SQLite, which have no Costura).
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
        Files = @('{0}.deps.json'); Copy = @{ 'RegistryExplorer\RegistryViewerZ\RegistryExplorerManual.pdf' = 'RegistryExplorerManual.pdf'; 'RegistryExplorerBookmarks\Common' = 'Bookmarks\Common' }
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

function Build-Program($p, $fw, [string]$buildDir) {
    $project = Join-Path $Root $p.Project
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

    foreach ($p in $selected) {
        # a script (Get-ZimmermanTools.ps1): copied as it is into <Out>\<Name>\script, signed there by --sign (the
        # repository's copy stays unsigned) and zipped at the top
        if ($p.Script) {
            Write-Host '== ' -ForegroundColor Red -NoNewline
            Write-Host "$($p.Name) script" -ForegroundColor Green
            $src = Join-Path $Root $p.Script
            $dest = Join-Path $Out "$($p.Name)\script"
            if (-not (Test-Path -LiteralPath $src)) {
                Write-Host "   $($p.Name): $src not found" -ForegroundColor Red
                $results.Add([pscustomobject]@{ Program = $p.Name; Framework = 'script'; Result = "FAILED: $($p.Script) not found" })
                continue
            }
            if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Recurse -Force }
            Copy-Into $src (Join-Path $dest (Split-Path $p.Script -Leaf))
            Write-Note "Copied $src"
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

            try {
                if (Test-Path $buildDir) { Remove-Item -LiteralPath $buildDir -Recurse -Force }
                if (Test-Path $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
                $buildDir = Build-Program $p $fw $buildDir
                New-Item -ItemType Directory -Force $target | Out-Null

                #(@(): a one-item array from an if is unrolled to a string, and += then appends to the string)
                $files = @(if ($fw.Name -eq 'net4' -or $p.SingleFile) { '{0}.exe' } else { '{0}.exe', '{0}.dll', '{0}.runtimeconfig.json' })
                $files += @($p.Files)
                foreach ($f in $files | Where-Object { $_ }) {
                    $name = $f -f $asm
                    $src = Join-Path $buildDir $name
                    if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination $target -Force }
                    elseif ($name -notlike '*.runtimeconfig.json' -or $fw.Name -ne 'net4') { throw "'$name' is not in the build output" }
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

                # every dll the build made must be in the program: staged next to it, or embedded by Costura (a build
                # without either can't run: the dependencies would be missing)
                $embedded = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                $main = Join-Path $target ($(if ($fw.Name -eq 'net4' -or $p.SingleFile) { '{0}.exe' } else { '{0}.dll' }) -f $asm)
                if (Test-Path -LiteralPath $main) { $embedded.UnionWith((Get-CosturaEmbedded $main)) }
                if ((Get-Content -LiteralPath $project -Raw) -match 'Costura\.Fody' -and $embedded.Count -eq 0) {
                    throw "$(Split-Path $main -Leaf) has no embedded dependencies (Costura did not run)"
                }

                #(staged anywhere in the program: Timeline Explorer's plugins are in Plugins)
                $stagedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                Get-ChildItem -LiteralPath $stage -Recurse -File -Filter *.dll | ForEach-Object { [void]$stagedNames.Add($_.Name) }
                $missing = @(Get-ChildItem $buildDir -Filter *.dll -File | Where-Object {
                    -not $stagedNames.Contains($_.Name) -and -not $embedded.Contains($_.Name) })
                if ($missing) {
                    throw "$($missing.Count) dll(s) of the build are neither embedded nor staged: $(($missing.Name | Select-Object -First 4) -join ', ')"
                }
                $note = if ($p.SingleFile) { 'staged (single file)' } else { 'staged' }
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

    # everything another machine needs to sign and zip: this script, and what was staged (with the commands)
    $self = Join-Path $Out (Split-Path $PSCommandPath -Leaf)
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
    $dir = [IO.Path]::GetFullPath($dir)
    Write-Host "Finding the .exe, .dll and .ps1 files in $dir"
    #(only the programs' folders: not the copy of this script at the top of a staged folder)
    $files = @(Get-ChildItem -LiteralPath $dir -Recurse -File -Include *.exe, *.dll, *.ps1 |
        Where-Object { $_.FullName -notmatch '\\_(build|release|zips)\\' -and
            -not ($_.Extension -eq '.ps1' -and $_.DirectoryName.TrimEnd('\') -ieq $dir.TrimEnd('\')) })
    $total = $files.Count
    Write-Note ("{0:N0} file(s), {1:N0} MB" -f $total, (($files | Measure-Object Length -Sum).Sum / 1MB))

    if (-not $Force -and $total -gt 0) {
        Write-Note "Checking their signatures (files already validly signed, such as DevExpress' and Microsoft's, are skipped; --force signs them too)"
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
