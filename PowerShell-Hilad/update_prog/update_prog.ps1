function UpdateProgramm {
    param(
        [string]$ProgramName,
        [string]$DownloadPath,
        [string]$DownloadProgram,
        [string]$UpdateDatei
    )

    # Erstelle die Quelle im Ereignisprotokoll, wenn sie nicht vorhanden ist
    $source = $programName
    if (-not ([System.Diagnostics.EventLog]::SourceExists($source))) {
        New-EventLog -LogName Application -Source $source
    }

    try {
        # Überprüfe, ob die heruntergeladene Datei existiert
        if (Test-Path $updateDatei) {
            # Ermittle die Version der heruntergeladenen Datei
            $newVersion = (Get-Item $updateDatei).VersionInfo.FileVersion
            
            # Überprüfe, ob das Programm installiert ist
            $installedProgram = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $programName }

            if ($installedProgram) {
                # Überprüfe, ob die installierte Version kleiner ist als die neue Version
                $installedVersion = $installedProgram.Version
                if ($installedVersion -lt $newVersion) {
                    # Wenn die installierte Version kleiner ist als die neue Version, führe das Update durch
                    # Stoppe das Programm, falls es läuft
                    Stop-Process -Name $programName -ErrorAction SilentlyContinue

                    # Deinstalliere das vorhandene Programm
                    $installedProgram.Uninstall()

                    # Installiere das Update, indem du die neue ausführbare Datei ausführst
                    Start-Process -FilePath $updateDatei -ArgumentList "/S" -Wait -Verb RunAs
                    $message = "Update fuer $programName auf Version $newVersion erfolgreich durchgefuehrt."
                    Remove-Item $updateDatei -Force
                    Write-Host $message
                    Write-EventLog -LogName Application -Source $programName -EntryType Information -EventId 1 -Message $message 
                } else {
                    $message = "Die installierte Version von $programName ist bereits auf dem neuesten Stand."
                    Write-Host $message
                    Write-EventLog -LogName Application -Source $programName -EntryType Information -EventId 2 -Message $message 
                }
            } else {
                # Wenn das Programm nicht installiert ist, führe die Installation der neuen Version durch
                Start-Process -FilePath $updateDatei -ArgumentList "/S" -Wait -Verb RunAs
                $message = "$programName wurde erfolgreich installiert in Version $newVersion."
                Write-Host $message
                Write-EventLog -LogName Application -Source $programName -EntryType Information -EventId 3 -Message $message 
            }
        } else {
            $message = "Die heruntergeladene Datei für $programName wurde nicht gefunden im Pfad: $updateDatei"
            Write-Host $message
            Write-EventLog -LogName Application -Source $programName -EntryType Error -EventId 4 -Message $message 
        }
    } catch {
        $errorMessage = "Ein Fehler ist aufgetreten: $($_.Exception.Message)"
        Write-Host $errorMessage
        Write-EventLog -LogName Application -Source $programName -EntryType Error -EventId 5 -Message $errorMessage 
    }
}

$downloadsPath = "C:\hilad"
$downloadProgram = "sipacompact.exe"
$programName = "Sicherheitspaket compact"

$timeout = New-TimeSpan -Minutes 5
$startTime = Get-Date

# $FileUri = "https://hilad.sharepoint.com/:u:/r/sites/EntwicklungfrMandanten/Freigegebene Dokumente/Download/" + $downloadProgram +"?csf=1&web=1&e=IPyDB8"
$FileUri = "https://download.datev.de/download/sipacompact/" + $downloadProgram +"?csf=1&web=1&e=IPyDB8"

$updateDatei = $downloadsPath + "\" + $downloadProgram

# Überprüfen, ob das Verzeichnis vorhanden ist
if (-not (Test-Path -Path $downloadsPath)) {
    # Verzeichnis erstellen, falls es nicht existiert
    New-Item -ItemType Directory -Path $downloadsPath
} 

try {

    Start-BitsTransfer -Source $FileUri -Destination $updateDatei

    # Warten, bis die Datei vollständig heruntergeladen ist oder der Timeout erreicht ist
    $startTime = Get-Date
    while ((Get-Date) -lt ($startTime + $timeout) -and -not (Test-Path $updateDatei)) {
        Start-Sleep -Seconds 1
    }

    # Überprüfen, ob die Datei erfolgreich heruntergeladen wurde
    if (Test-Path $updateDatei) {
        Write-Host "Die Datei wurde erfolgreich heruntergeladen und ist nun verfügbar."
        UpdateProgramm -ProgramName $programName -DownloadPath $downloadsPath -DownloadProgram $downloadProgram -UpdateDatei $updateDatei
    } else {
        Write-Host "Der Download wurde nicht innerhalb des Timeouts abgeschlossen. Vorgang wird abgebrochen."
    }
} catch {
    Write-Host "Beim Herunterladen der Datei ist ein Fehler aufgetreten: $_"
} finally {
   
}
