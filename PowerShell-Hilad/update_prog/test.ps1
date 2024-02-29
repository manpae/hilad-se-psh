$downloadsPath = "C:\hilad"
$downloadProgram = "sipacompact.exe"
$FileUri = "https://hilad.sharepoint.com/:u:/r/sites/EntwicklungfrMandanten/Freigegebene Dokumente/Download/" + $downloadProgram +"?csf=1&web=1&e=IPyDB8"
$FileUri = "https://download.datev.de/download/sipacompact/sipacompact.exe"
Write-Host $FileUri
$updateDatei = $downloadsPath + "\" + $downloadProgram

# Invoke-WebRequest -URI $FileUri -OutFile $updateDatei

#(New-Object System.Net.WebClient).DownloadFile($FileUri, $updateDatei)

Start-BitsTransfer -Source $FileUri -Destination $updateDatei

