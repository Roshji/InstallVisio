<# ###################
Author                  :M van Rijn - Prodicom
Updated for Win11 23h2  :R van Diermen - RvD IT Consulting
 
Uninstall $Productname based on XML
#################### #>
 
# Vars
$Date = (Get-Date).tostring("yyyy-MM-dd HHmm")
$Global:ProductName = "Visio365"
$Global:ExtractAppFolder = "Officedeploymenttool"
$Global:ODTExe = "officedeploymenttool.exe"
$Global:DownloadRootFolder = "C:\ProgramData\Intune\Packages\$($Global:ProductName)"
$LogFile = "C:\Programdata\Microsoft\IntuneManagementExtension\Logs\Logs_$($Global:ProductName)_uninstall_$($Date).log"
$Global:installFolder = "$PSScriptRoot"
$Url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
$Response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
$Global:SetupEXEFile = "$Global:DownloadRootFolder\$Global:ExtractAppFolder\Setup.exe"
$Global:UnInstallXML = "Uninstall_$Global:ProductName.xml"
 
# Start Logging
Start-Transcript -Path $LogFile
 
# Check if original Setup files are present on the system, if not download ODT
If (!(Test-Path $Global:SetupEXEFile)) {
    Write-host "Setup.exe for uninstall not found. Download ODT package"
    ###########################
    # FOLDER CREATION SECTION #
    ###########################
    # Check if folders exist else create
    If (!(Test-Path $Global:DownloadRootFolder)) {
        # Create directory
        New-item -path "$Global:DownloadRootFolder" -ItemType "directory"
    }
    If (!(Test-Path $Global:DownloadRootFolder)) {
        # Create directory
        New-item -path "$Global:DownloadRootFolder" -ItemType "directory"
    }
 
    ####################
    # DOWNLOAD SECTION #
    ####################
    # Get Download URL of latest Office Deployment Tool (ODT)
    $ODTUri = $response.links | Where-Object {$_.outerHTML -like "*click here to download manually*"}
    $UrlCurrentVerODT = $ODTUri.href
    # Download latest Office Deployment Tool (ODT)
    Write-Host "Downloading latest version of Office Deployment Tool (ODT)."
    Invoke-WebRequest -UseBasicParsing -Uri $UrlCurrentVerODT -OutFile $Global:DownloadRootFolder\$Global:ODTExe
    # Get file version
    Write-Host "Get fileversion Office Deployment Tool (ODT)."
    $Version = (Get-Command $Global:DownloadRootFolder\$Global:ODTExe).FileVersionInfo.FileVersion
    Write-Host "Fileversion Office Deployment Tool (ODT): " $Version
    # Unpack ODT File
    Write-Host "Unpacking file ..."
    $arguments1 = "/quiet /extract:$Global:DownloadRootFolder\$Global:ExtractAppFolder"
    Start-Process -Wait -FilePath "$Global:DownloadRootFolder\$Global:ODTExe" -ArgumentList $arguments1
    Start-sleep -s 5
}
 
# Remove
Write-Host "Remove $Global:ProductName..."
$arguments1 = "/Configure $Global:installFolder\$Global:UnInstallXML"
Start-Process -Wait -FilePath "$Global:SetupEXEFile" -ArgumentList $arguments1
 
###################
# CLEANUP SECTION #
###################
Start-sleep -s 5
# Cleanup
Write-Host "Cleanup, remove folder $Global:DownloadRootFolder "
Remove-Item -Path "$Global:DownloadRootFolder" -Recurse -Force
 
# Stop Logging
Stop-Transcript