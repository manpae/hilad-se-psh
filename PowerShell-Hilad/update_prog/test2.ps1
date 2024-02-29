# Setze das Standard-Download-Verzeichnis
[Environment]::SetEnvironmentVariable("USERPROFILE", $env:USERPROFILE, [System.EnvironmentVariableTarget]::User)
$downloadsPath = Join-Path $env:USERPROFILE "Downloads"

# Definiere den Namen des Programms und den Pfad zur neuen ausführbaren Datei im Benutzer-Download-Verzeichnis
$programName = "Sicherheitspaket compact"
$newExecutablePath = Join-Path -Path $downloadsPath -ChildPath "sipacompact.exe"
$newVersion = "7.6.104.24048" # (Get-Item $newExecutablePath).VersionInfo.FileVersion

# Erstelle die Quelle im Ereignisprotokoll, wenn sie nicht vorhanden ist
$source = $programName
if (-not ([System.Diagnostics.EventLog]::SourceExists($source))) {
    New-EventLog -LogName Application -Source $source
}

try {
    # Überprüfe, ob die heruntergeladene Datei existiert
    if (Test-Path $newExecutablePath) {
        # Ermittle die Version der heruntergeladenen Datei
        $newVersion = (Get-Item $newExecutablePath).VersionInfo.FileVersion
        
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
                Start-Process -FilePath $newExecutablePath -ArgumentList "/S" -Wait -Verb RunAs
                $message = "Update fuer $programName auf Version $newVersion erfolgreich durchgefuehrt."
                Write-Host $message
                Write-EventLog -LogName Application -Source $programName -EntryType Information -EventId 1 -Message $message 
            } else {
                $message = "Die installierte Version von $programName ist bereits auf dem neuesten Stand."
                Write-Host $message
                Write-EventLog -LogName Application -Source $programName -EntryType Information -EventId 2 -Message $message 
            }
        } else {
            # Wenn das Programm nicht installiert ist, führe die Installation der neuen Version durch
            Start-Process -FilePath $newExecutablePath -ArgumentList "/S" -Wait -Verb RunAs
            $message = "$programName wurde erfolgreich installiert in Version $newVersion."
            Write-Host $message
            Write-EventLog -LogName Application -Source $programName -EntryType Information -EventId 3 -Message $message 
        }
    } else {
        $message = "Die heruntergeladene Datei für $programName wurde nicht gefunden im Pfad: $userDownloadPath"
        Write-Host $message
        Write-EventLog -LogName Application -Source $programName -EntryType Error -EventId 4 -Message $message 
    }
} catch {
    $errorMessage = "Ein Fehler ist aufgetreten: $($_.Exception.Message)"
    Write-Host $errorMessage
    Write-EventLog -LogName Application -Source $programName -EntryType Error -EventId 5 -Message $errorMessage 
}
