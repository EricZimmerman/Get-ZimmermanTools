[Cmdletbinding()]
# Where to extract the files to
Param
(
    $Dest
)

Write-Host "`nThs script will discover and download all available programs from https://ericzimmerman.github.io and download them to $Dest" -BackgroundColor Blue
Write-Host "A file will also be created in $Dest that tracks the SHA-1 of each file, so rerunning the script will only download new versions"
Write-Host "To redownload, remove lines from or delete the CSV file created under $Dest and rerun. Enjoy!`n"

$newInstall = $false

if(!(Test-Path -Path $Dest ))
{
    write-host $Dest + " does not exist. Creating..."
    New-Item -ItemType directory -Path $Dest > $null

    $newInstall = $true
}

$URL = "https://raw.githubusercontent.com/EricZimmerman/ericzimmerman.github.io/master/index.md"


$WebKeyCollection = @()


$localDetailsFile = Join-Path $Dest -ChildPath "!!!RemoteFileDetails.csv"

if (Test-Path -Path $localDetailsFile)
{
    write-host "Loading local details..."
    $LocalKeyCollection = Import-Csv -Path $localDetailsFile
}

$toDownload = @()

#Get zips
$PageContent = (Invoke-WebRequest -Uri $URL).Content

$regex = [regex] '(?i)\b(https)://[-A-Z0-9+&@#/%?=~_|$!:,.;]*[A-Z0-9+&@#/%=~_|$].zip'
$matchdetails = $regex.Match($PageContent)
while ($matchdetails.Success) {
    $headers = (Invoke-WebRequest -Uri $matchdetails.Value -Method Head).Headers

    $getUrl = $matchdetails.Value
    $sha = $headers["x-bz-content-sha1"]
    $name = $headers["x-bz-file-name"]

    $details = @{            
        Name     = $name            
        SHA1     = $sha                 
        URL     = $getUrl
        }                           

    $webKeyCollection += New-Object PSObject -Property $details  

	$matchdetails = $matchdetails.NextMatch()
} 

Foreach ($webKey in $webKeyCollection)
{
    if ($newInstall)
    {
        $toDownload+= $webKey
        continue    
    }

    $localFile = $LocalKeyCollection | Where-Object {$_.Name -eq $webKey.Name}

    if ($null -eq $localFile -or $localFile.SHA1 -ne $webKey.SHA1)
    {
        $toDownload+= $webKey
    }
}

if ($toDownload.Count -eq 0)
{
    write-host "`nAll files current. Exiting.`n" -BackgroundColor Blue
    return
}

if (-not (test-path ".\7z\7za.exe")) 
{
    Write-Host "`n.\7z\7za.exe needed! Exiting`n" -BackgroundColor Red
    return
} 
set-alias sz ".\7z\7za.exe"  

foreach($td in $toDownload)
{
    $dUrl = $td.URL
    write-host "Downloading $dUrl" -ForegroundColor Green
    $destFile = Join-Path -Path . -ChildPath $td.Name
    Invoke-WebRequest -Uri $dUrl -OutFile $destFile

    write-host "`tUnzipping to $Dest..."
    sz x $destFile -o"$Dest" -y > $null
    remove-item -Path $destFile
}

Write-host "`nSaving version information to $localDetailsFile`n" -ForegroundColor Red
$webKeyCollection | export-csv -Path  $localDetailsFile
