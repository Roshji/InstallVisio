# Check if App is installed

$AppFolder = "C:\Program Files\Microsoft Office\root\Office16"
$AppFile = "VISIO.EXE"

If (Test-Path $AppFolder\$AppFile) {
    Write-Output "App detected"
    Exit 0
} Else {
    Exit 1
}