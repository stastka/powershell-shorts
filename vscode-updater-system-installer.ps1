<#
.SYNOPSIS
	Update vscode over Powershell
.DESCRIPTION
	This PowerShell script update vscode with system installer over Powershell
.EXAMPLE
	PS> ./vscode-updater-system-installer.ps1
.LINK
	https://github.com/stastka/powershell-shorts
.NOTES
	Author: Daniel Stastka | License: Apache2
  Tested with w2k19,w2k22,w2k25,win10
#>

$winversion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion") | Select-Object -Property ProductName,CurrentBuild,ReleaseId
if($winversion.ReleaseId -eq 1809)
{
    #17763,1809 = MS Windows Server 2019
    #20348,2009 = MS Windows Server 2022
    #26100,2009 = MS Windows Server 2025
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -Value 2
}
$localRepo = "C:\ProgramData\repo"

if((Test-Path -Path $localRepo) -eq $false)
{
    $localRepo = New-item -Path "$localRepo" -ItemType Directory -ErrorAction SilentlyContinue -Force
}

$dirLocalRoot = Get-Item -Path $localRepo
$dirDownload = New-item -Path "$dirLocalRoot\vscode" -ItemType Directory -ErrorAction SilentlyContinue -Force
$versionBaseName = "VSCode-"

$downloadVersionUri = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64"
$feedURI = "https://code.visualstudio.com/feed.xml"

$installedVersionFile = (Get-Item "C:\Program Files\Microsoft VS Code\Code.exe").VersionInfo.FileVersionRaw
$installedVersion = "$($installedVersionFile.Major).$($installedVersionFile.Minor)"

$feedPosts = Invoke-RestMethod -Uri $feedURI

# search for new versions in feed
$feedVersions = @()
$feedPosts | Select-Object -First 5 | ForEach-Object {
    if ($_.Id -match "updates/v[0-9]{1}_[0-9]{2}") {

        # extract version from regex match
        $feedVersion = ($Matches[0] -split "updates/v")[1] -replace "_","."

        # expand version to full release number
        if ($feedVersion.split(".").count -eq 2) {
            $feedVersion = $feedVersion + ".0"
        }

        # object to store version
        $tempVersion = [PSCustomObject]@{
            Full = [string]$feedVersion
            Major = [decimal](($feedVersion.Split('.',3) | Select-Object -Index 0,1) -join ".")
            Minor = [decimal](($feedVersion.Split('.',3) | Select-Object -Index 1,2) -join ".")
        }

        # update collection array
        $feedVersions += $tempVersion
    }
}
$feedLatestVersion = $feedVersions | Sort-Object Full -Descending | Select-Object -First 1

if($installedVersion -eq $feedLatestVersion.Major)
{
    Write-Host "VSCode Installed: $($installedVersion) ($("$($installedVersionFile.Major).$($installedVersionFile.Minor).$($installedVersionFile.Build)"))" -ForegroundColor Green
    Write-Host "VSCode Feed:      $($feedLatestVersion.Major) ($($feedLatestVersion.Full))" -ForegroundColor Green
}
else {
    Write-Host "VSCode Installed: $($installedVersion)" -ForegroundColor Red
    Write-Host "VSCode Feed: $($feedLatestVersion.Major)" -ForegroundColor Red

    $downloadFilePath = "$($dirDownload)\$($versionBaseName)$($feedLatestVersion.Full).exe"
    if((Test-Path -Path $downloadFilePath) -eq $false)
    {
        $downloadRequest = Invoke-WebRequest -Uri $downloadVersionUri -OutFile $downloadFilePath -PassThru
        # check HTTP StatusCode
        if ($downloadRequest.StatusCode -ne 200) {
            Write-Verbose "Download failed with error '$($downloadRequest.StatusCode) - $($downloadRequest.StatusDescription)'" -Verbose
            Remove-Item -Path $downloadFilePath -Force -Verbose -ErrorAction SilentlyContinue
            exit
        }
    }
    else {
        Write-Host "$downloadFilePath Existis" -ForegroundColor Green
    }
    if((Test-Path -Path $downloadFilePath) -eq $true)
    {
        $logfile = "$($localRepo)\vscode\vscode_install$($feedLatestVersion.Full).log"
        Start-Process -FilePath $downloadFilePath -Argument "/VERYSILENT /CLOSEAPPLICATIONS /MERGETASKS=!runcode /LOG=`"$logfile`""
    }
}
