<# ###################
Author                  :M van Rijn - Prodicom
Updated for Win11 23h2  :R van Diermen - RvD IT Consulting
 
# --- INSTALL ---
# Specify languages in "Install_[Productname].xml" for installation
 
- Create download folder for Officedeploymenttool
- Download latest version of the Officedeploymenttool
- Unpack Officedeploymenttool
- Download [Productname] via Officedeploymenttool and XML to the (package) $installfolder, which will be removed after installation
- Install [Productname] via Officedeploymenttool and XML
 
The download folder will remain on the device for uninstall
#################### #>
 
# Vars
$Date = (Get-Date).tostring("yyyy-MM-dd HHmm")
$Global:ProductName = "Visio365"
$Global:ExtractAppFolder = "Officedeploymenttool"
$Global:ODTExe = "officedeploymenttool.exe"
$Global:DownloadAppFolder = "C:\ProgramData\Intune\Packages\$($Global:ProductName)"
$LogFile = "C:\Programdata\Microsoft\IntuneManagementExtension\Logs\Logs_$($Global:ProductName)_install_$($Date).log"
$Global:installFolder = "$PSScriptRoot"
$Url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
$Response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
$Global:SetupEXEFile = "$Global:DownloadAppFolder\$Global:ExtractAppFolder\Setup.exe"
$Global:InstallXML = "Install_$Global:ProductName.xml"
 
# Start Logging
Start-Transcript -Path $LogFile
 
###########################
# FOLDER CREATION SECTION #
###########################
If (!(Test-Path $Global:DownloadAppFolder)) {
    # Create directory
    New-item -path "$Global:DownloadAppFolder" -ItemType "directory"
}
####################
# DOWNLOAD SECTION #
####################
# Get Download URL of latest Office Deployment Tool (ODT)
$ODTUri = $response.links | Where-Object {$_.outerHTML -like "*click here to download manually*"}
$UrlCurrentVerODT = $ODTUri.href
# Download latest Office Deployment Tool (ODT)
Write-Host "Downloading latest version of Office Deployment Tool (ODT)."
Invoke-WebRequest -UseBasicParsing -Uri $UrlCurrentVerODT -OutFile $Global:DownloadAppFolder\$Global:ODTExe
# Get file version
Write-Host "Get fileversion Office Deployment Tool (ODT)."
$Version = (Get-Command $Global:DownloadAppFolder\$Global:ODTExe).FileVersionInfo.FileVersion
Write-Host "Fileversion Office Deployment Tool (ODT): " $Version
# Unpack ODT File
Write-Host "Unpacking file ..."
$arguments1 = "/quiet /extract:$Global:DownloadAppFolder\$Global:ExtractAppFolder"
Start-Process -Wait -FilePath "$Global:DownloadAppFolder\$Global:ODTExe" -ArgumentList $arguments1
Start-sleep -s 5
 
# Download
Write-Host "Download $Global:ProductName..."
$arguments1 = "/download $Global:installFolder\$Global:InstallXML"
Start-Process -Wait -FilePath "$Global:SetupEXEFile" -ArgumentList $arguments1
 
# Install
Write-Host "Installing $Global:ProductName..."
$arguments1 = "/configure $Global:installFolder\$Global:InstallXML"
Start-Process -Wait -FilePath "$Global:SetupEXEFile" -ArgumentList $arguments1
 
###################
# CLEANUP SECTION #
###################
Start-sleep -s 20
# Cleanup ODT file. Folder with officedeploymenttool and software will remain for uninstall. Change if needed
Write-Host "Cleanup ODT download file"
Remove-Item "$Global:DownloadAppFolder\officedeploymenttool.exe"
 
# Stop Logging
Stop-Transcript