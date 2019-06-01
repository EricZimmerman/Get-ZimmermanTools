<#
.SYNOPSIS
    This script will discover and download all available programs from https://ericzimmerman.github.io and download them to $Dest
.DESCRIPTION
    A file will also be created in $Dest that tracks the SHA-1 of each file, so rerunning the script will only download new versions. To redownload, remove lines from or delete the CSV file created under $Dest and rerun.
.PARAMETER Dest
    The path you want to save the programs to.
.EXAMPLE
    C:\PS> Get-ZimmermanTools.ps1 -Dest c:\tools
    Downloads/extracts and saves details about programs to c:\tools directory.
.NOTES
    Author: Eric Zimmerman
    Date:   January 22, 2019    
#>

[Cmdletbinding()]
# Where to extract the files to
Param
(
    [Parameter()]
    [string]$Dest = (Resolve-Path ".") #Where to save programs to	
)


function Write-Color {
    <#
	.SYNOPSIS
        Write-Color is a wrapper around Write-Host.
        It provides:
        - Easy manipulation of colors,
        - Logging output to file (log)
        - Nice formatting options out of the box.
	.DESCRIPTION
        Author: przemyslaw.klys at evotec.pl
        Project website: https://evotec.xyz/hub/scripts/write-color-ps1/
        Project support: https://github.com/EvotecIT/PSWriteColor
        Original idea: Josh (https://stackoverflow.com/users/81769/josh)
	.EXAMPLE
    Write-Color -Text "Red ", "Green ", "Yellow " -Color Red,Green,Yellow
    .EXAMPLE
	Write-Color -Text "This is text in Green ",
					"followed by red ",
					"and then we have Magenta... ",
					"isn't it fun? ",
					"Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan
    .EXAMPLE
	Write-Color -Text "This is text in Green ",
					"followed by red ",
					"and then we have Magenta... ",
					"isn't it fun? ",
                    "Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan -StartTab 3 -LinesBefore 1 -LinesAfter 1
    .EXAMPLE
	Write-Color "1. ", "Option 1" -Color Yellow, Green
	Write-Color "2. ", "Option 2" -Color Yellow, Green
	Write-Color "3. ", "Option 3" -Color Yellow, Green
	Write-Color "4. ", "Option 4" -Color Yellow, Green
	Write-Color "9. ", "Press 9 to exit" -Color Yellow, Gray -LinesBefore 1
    .EXAMPLE
	Write-Color -LinesBefore 2 -Text "This little ","message is ", "written to log ", "file as well." `
				-Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt" -TimeFormat "yyyy-MM-dd HH:mm:ss"
	Write-Color -Text "This can get ","handy if ", "want to display things, and log actions to file ", "at the same time." `
				-Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt"
    .EXAMPLE
    # Added in 0.5
    Write-Color -T "My text", " is ", "all colorful" -C Yellow, Red, Green -B Green, Green, Yellow
    wc -t "my text" -c yellow -b green
    wc -text "my text" -c red
    .NOTES
        CHANGELOG
        Version 0.5 (25th April 2018)
        -----------
        - Added backgroundcolor
        - Added aliases T/B/C to shorter code
        - Added alias to function (can be used with "WC")
        - Fixes to module publishing
        Version 0.4.0-0.4.9 (25th April 2018)
        -------------------
        - Published as module
        - Fixed small issues
        Version 0.31 (20th April 2018)
        ------------
        - Added Try/Catch for Write-Output (might need some additional work)
        - Small change to parameters
        Version 0.3 (9th April 2018)
        -----------
        - Added -ShowTime
        - Added -NoNewLine
        - Added function description
        - Changed some formatting
        Version 0.2
        -----------
        - Added logging to file
        Version 0.1
        -----------
        - First draft
        Additional Notes:
        - TimeFormat https://msdn.microsoft.com/en-us/library/8kb3ddd4.aspx
    #>
    [alias('Write-Colour')]
    [CmdletBinding()]
    param (
        [alias ('T')] [String[]]$Text,
        [alias ('C', 'ForegroundColor', 'FGC')] [ConsoleColor[]]$Color = [ConsoleColor]::White,
        [alias ('B', 'BGC')] [ConsoleColor[]]$BackGroundColor = $null,
        [alias ('Indent')][int] $StartTab = 0,
        [int] $LinesBefore = 0,
        [int] $LinesAfter = 0,
        [int] $StartSpaces = 0,
        [alias ('L')] [string] $LogFile = '',
        [Alias('DateFormat', 'TimeFormat')][string] $DateTimeFormat = 'yyyy-MM-dd HH:mm:ss',
        [alias ('LogTimeStamp')][bool] $LogTime = $true,
        [ValidateSet('unknown', 'string', 'unicode', 'bigendianunicode', 'utf8', 'utf7', 'utf32', 'ascii', 'default', 'oem')][string]$Encoding = 'Unicode',
        [switch] $ShowTime,
        [switch] $NoNewLine
    )
    $DefaultColor = $Color[0]
    if ($null -ne $BackGroundColor -and $BackGroundColor.Count -ne $Color.Count) { Write-Error "Colors, BackGroundColors parameters count doesn't match. Terminated." ; return }
    #if ($Text.Count -eq 0) { return }
    if ($LinesBefore -ne 0) { for ($i = 0; $i -lt $LinesBefore; $i++) { Write-Host -Object "`n" -NoNewline } } # Add empty line before
    if ($StartTab -ne 0) { for ($i = 0; $i -lt $StartTab; $i++) { Write-Host -Object "`t" -NoNewLine } }  # Add TABS before text
    if ($StartSpaces -ne 0) { for ($i = 0; $i -lt $StartSpaces; $i++) { Write-Host -Object ' ' -NoNewLine } }  # Add SPACES before text
    if ($ShowTime) { Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))]" -NoNewline } # Add Time before output
    if ($Text.Count -ne 0) {
        if ($Color.Count -ge $Text.Count) {
            # the real deal coloring
            if ($null -eq $BackGroundColor) {
                for ($i = 0; $i -lt $Text.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewLine }
            }
            else {
                for ($i = 0; $i -lt $Text.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewLine }
            }
        }
        else {
            if ($null -eq $BackGroundColor) {
                for ($i = 0; $i -lt $Color.Length ; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewLine }
                for ($i = $Color.Length; $i -lt $Text.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -NoNewLine }
            }
            else {
                for ($i = 0; $i -lt $Color.Length ; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewLine }
                for ($i = $Color.Length; $i -lt $Text.Length; $i++) { Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $BackGroundColor[0] -NoNewLine }
            }
        }
    }
    if ($NoNewLine -eq $true) { Write-Host -NoNewline } else { Write-Host } # Support for no new line
    if ($LinesAfter -ne 0) { for ($i = 0; $i -lt $LinesAfter; $i++) { Write-Host -Object "`n" -NoNewline } }  # Add empty line after
    if ($Text.Count -ne 0 -and $LogFile -ne "") {
        # Save to file
        $TextToFile = ""
        for ($i = 0; $i -lt $Text.Length; $i++) {
            $TextToFile += $Text[$i]
        }
        try {
            if ($LogTime) {
                Write-Output -InputObject "[$([datetime]::Now.ToString($DateTimeFormat))]$TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append
            }
            else {
                Write-Output -InputObject "$TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append
            }
        }
        catch {
            $_.Exception
        }
    }
}

Write-Host ""
Write-Host "This script will discover and download all available programs" -BackgroundColor Blue
Write-Host "from https://ericzimmerman.github.io and download them to $Dest" -BackgroundColor Blue
Write-Host "`nA file will also be created in $Dest that tracks the SHA-1 of each file,"
Write-Host "so rerunning the script will only download new versions."
Write-Host "`nTo redownload, remove lines from or delete the CSV file created under $Dest and rerun. Enjoy!`n"

$defaultColor = (get-host).ui.rawui.ForegroundColor

$newInstall = $false

if (!(Test-Path -Path $Dest )) {
    write-host $Dest " does not exist. Creating..."
    New-Item -ItemType directory -Path $Dest > $null

    $newInstall = $true
}

$URL = "https://raw.githubusercontent.com/EricZimmerman/ericzimmerman.github.io/master/index.md"

$WebKeyCollection = @()

$localDetailsFile = Join-Path $Dest -ChildPath "!!!RemoteFileDetails.csv"

if (Test-Path -Path $localDetailsFile) {
    write-color -Text "* ", "Loading local details from '$Dest'..." -Color Green, $defaultColor
    $LocalKeyCollection = Import-Csv -Path $localDetailsFile
}

$toDownload = @()

#Get zips
$progressPreference = 'silentlyContinue'
$PageContent = (Invoke-WebRequest -Uri $URL -UseBasicParsing).Content
$progressPreference = 'Continue'

$regex = [regex] '(?i)\b(https)://[-A-Z0-9+&@#/%?=~_|$!:,.;]*[A-Z0-9+&@#/%=~_|$].(zip|txt)'
$matchdetails = $regex.Match($PageContent)

write-color -Text "* ", "Getting available programs..." -Color Green, $defaultColor
$progressPreference = 'silentlyContinue'
while ($matchdetails.Success) {
    $headers = (Invoke-WebRequest -Uri $matchdetails.Value -UseBasicParsing -Method Head).Headers

    $getUrl = $matchdetails.Value
    $sha = $headers["x-bz-content-sha1"]
    $name = $headers["x-bz-file-name"]
    $size = $headers["Content-Length"]

    $details = @{            
        Name = [string]$name            
        SHA1 = [string]$sha                 
        URL  = [string]$getUrl
        Size = [string]$size
    }                           

    $webKeyCollection += New-Object PSObject -Property $details  

    $matchdetails = $matchdetails.NextMatch()
} 
$progressPreference = 'Continue'

Foreach ($webKey in $webKeyCollection) {
    if ($newInstall) {
        $toDownload += $webKey
        continue    
    }

    $localFile = $LocalKeyCollection | Where-Object { $_.Name -eq $webKey.Name }

    if ($null -eq $localFile -or $localFile.SHA1 -ne $webKey.SHA1) {
        #Needs to be downloaded since SHA is different or it doesnt exist
        $toDownload += $webKey
    }
}

if ($toDownload.Count -eq 0) {
    Write-Host ""
  
    write-color -Text "* ", "All files current. Exiting." -Color Green, Blue
    Write-Host "`n"
    return
}

if (-not (test-path ".\7z\7za.exe")) {
    Write-Host "`n.\7z\7za.exe needed! Exiting`n" -BackgroundColor Red
    return
} 
set-alias sz ".\7z\7za.exe"  

$downloadedOK = @()

$destFile = ""
$name = ""

$i = 0
$dlCount = $toDownload.Count
write-color -Text "* ", "Files to download: $dlCount" -Color Green, $defaultColor
foreach ($td in $toDownload) {
    $p = [math]::round( ($i / $toDownload.Count) * 100, 2 )

    #Write-Host ($td | Format-Table | Out-String)
    
    try {
        $dUrl = $td.URL
        $size = $td.Size
        $name = $td.Name

        Write-Progress -Activity "Updating programs...." -Status "$p% Complete" -PercentComplete $p -CurrentOperation "Downloading $name" 
        $destFile = [IO.Path]::Combine(".", $name)

        $progressPreference = 'silentlyContinue'
        Invoke-WebRequest -Uri $dUrl -OutFile $destFile -ErrorAction:Stop -UseBasicParsing

        write-color -Text "* ", "Downloaded $name (Size: $size)" -Color Green, Blue
    
        $downloadedOK += $td

        if ( $name.endswith("zip") ) {
            sz x $destFile -o"$Dest" -y > $null
        }      
    }
    catch {
        $ErrorMessage = $_.Exception.Message
        write-host "Error downloading $name ($ErrorMessage). Wait for the run to finish and try again by repeating the command"
    }
    finally {
        $progressPreference = 'Continue'
        if ( $name.endswith("zip") ) {
            remove-item -Path $destFile
        } 
        
    }
    $i += 1
}

#Write-Host ($webKeyCollection | Format-Table | Out-String)

#Downloaded ok contains new stuff, but we need to account for existing stuff too
foreach ($webItems in $webKeyCollection) {
    #Check what we have locally to see if it also contains what is in the web collection
    $localFile = $LocalKeyCollection | Where-Object { $_.SHA1 -eq $webItems.SHA1 }

    #if its not null, we have a local file match against what is on the website, so its ok
    
    if ($null -ne $localFile) {
        #consider it downloaded since SHAs match
        $downloadedOK += $webItems
    }
}


Write-Host ""
write-color -Text "* ", "Saving downloaded version information to $localDetailsFile" -Color Green, Red
Write-host "`n"
$downloadedOK | export-csv -Path  $localDetailsFile

