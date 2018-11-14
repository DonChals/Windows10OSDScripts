Param(
    [String]$CompanyName = ""
)

$LogPath = "$env:TEMP\CorrectTaskBar.log"
Start-Transcript -Path $LogPath

$XMLPath = $env:USERPROFILE + "\Appdata\local\microsoft\windows\shell\layoutmodification.xml"

Write-Host "Overwritting default start layout.."
Copy-Item C:\Windows\$CompanyName\LayoutModification.xml $XMLPath -Force -Verbose

Stop-Transcript