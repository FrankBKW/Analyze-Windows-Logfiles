# ============================================================
#  Windows Event Analyzer – Interaktives Abfrage-Tool
#  Anforderungen: Windows PowerShell 5.1 oder PowerShell 7+
#  Als Administrator ausführen für vollständigen Log-Zugriff
# ============================================================
#
# ⚠️  DISCLAIMER / RECHTLICHER HINWEIS
# ============================================================
#
#  NUTZUNGSZWECK
#  Dieses Script ist ausschließlich für den legitimen, autorisierten
#  Einsatz durch IT-Administratoren, System-Engineers und Sicherheits-
#  verantwortliche bestimmt, die zur Analyse der betroffenen Systeme
#  ausdrücklich berechtigt sind.
#
#  ZUGRIFF AUF FREMDE SYSTEME
#  Das Auslesen von Ereignisprotokollen ohne ausdrückliche Genehmigung
#  des Systemeigentümers oder zuständigen IT-Verantwortlichen kann
#  strafbar sein:
#    - Deutschland : § 202a StGB (Ausspähen von Daten)
#    - Österreich  : § 118a öStGB
#    - Schweiz     : Art. 143 StGB
#    - EU          : Richtlinie 2013/40/EU
#    - USA         : Computer Fraud and Abuse Act (CFAA)
#  Nur Systeme abfragen, für die eine ausdrückliche Berechtigung vorliegt.
#
#  DATENSCHUTZ / DSGVO
#  Windows-Ereignisprotokolle können personenbezogene Daten enthalten
#  (Benutzernamen, IP-Adressen, Anmeldezeiten, Dateizugriffe u.a.).
#  Die Verarbeitung unterliegt der DSGVO sowie nationalen Datenschutz-
#  gesetzen. Exportierte Dateien (CSV, Excel) sind entsprechend zu
#  schützen und nach Ablauf der Aufbewahrungsfrist sicher zu löschen.
#
#  CREDENTIALS
#  Eingegebene Passwörter werden als SecureString im Arbeitsspeicher
#  verwaltet und nicht dauerhaft gespeichert. Die Verantwortung für
#  den sicheren Umgang mit Zugangsdaten liegt beim Anwender.
#
#  HAFTUNGSAUSSCHLUSS
#  Dieses Script wird OHNE JEDE GEWÄHRLEISTUNG bereitgestellt.
#  Der Autor übernimmt keine Haftung für Schäden, die durch die
#  Nutzung, Fehlnutzung oder den Ausfall des Scripts entstehen,
#  einschließlich Datenverlust, Systemausfälle oder Sicherheits-
#  vorfälle. Der Einsatz erfolgt auf eigene Verantwortung.
#
#  EXECUTIONPOLICY
#  Der Start mit -ExecutionPolicy Bypass deaktiviert Skript-
#  signaturprüfungen. Für den Produktiveinsatz empfohlen:
#    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
#  oder das Script mit einem Code-Signing-Zertifikat signieren.
#
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization -ErrorAction SilentlyContinue

# ── Profil-Verzeichnis ────────────────────────────────────────
$script:profileDir = Join-Path $env:APPDATA "WindowsEventAnalyzer\profiles"
if (-not (Test-Path $script:profileDir)) {
    New-Item -ItemType Directory -Path $script:profileDir -Force | Out-Null
}

# ── Farbschema ───────────────────────────────────────────────
$clrBg        = [System.Drawing.Color]::FromArgb(245, 245, 250)
$clrPanel     = [System.Drawing.Color]::FromArgb(255, 255, 255)
$clrAccent    = [System.Drawing.Color]::FromArgb(74,  74, 170)
$clrAccentHov = [System.Drawing.Color]::FromArgb(94,  94, 190)
$clrBorder    = [System.Drawing.Color]::FromArgb(210, 210, 230)
$clrText      = [System.Drawing.Color]::FromArgb(30,  30,  50)
$clrMuted     = [System.Drawing.Color]::FromArgb(110, 110, 140)
$clrWarn      = [System.Drawing.Color]::FromArgb(200,  80,  60)
$clrSuccess   = [System.Drawing.Color]::FromArgb(40,  160,  80)
$clrInfo      = [System.Drawing.Color]::FromArgb(30,  120, 200)

$fontNormal   = New-Object System.Drawing.Font("Segoe UI", 9)
$fontBold     = New-Object System.Drawing.Font("Segoe UI", 9,  [System.Drawing.FontStyle]::Bold)
$fontSmall    = New-Object System.Drawing.Font("Segoe UI", 8)
$fontTitle    = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$fontMono     = New-Object System.Drawing.Font("Consolas",  8)

# ── Event-ID Katalog ─────────────────────────────────────────
# Format: ID | Protokoll (Log-Name) | Kategorie | Icon | Beschreibung
# Logs: Klassisch = "System", "Security", "Application"
#       Modern    = "Microsoft-Windows-.../..."  -> benötigen Get-WinEvent
$eventCatalog = [System.Collections.Generic.List[PSObject]]::new()

# ─── System-Events ────────────────────────────────────────────
$sysEvents = @(
    @{ ID=41;    Cat="Kritisch";    Icon="●"; Desc="Kernel-Power: Unerwarteter Neustart / Absturz" }
    @{ ID=55;    Cat="Kritisch";    Icon="●"; Desc="NTFS-Dateisystemfehler (Korruption erkannt)" }
    @{ ID=11;    Cat="Fehler";      Icon="●"; Desc="Treiber-Controllerfehler (Festplatte/USB)" }
    @{ ID=1074;  Cat="Warnung";     Icon="▲"; Desc="Geplanter Neustart/Shutdown durch Benutzer/Prozess" }
    @{ ID=6005;  Cat="Information"; Icon="○"; Desc="Ereignisprotokoll-Dienst gestartet (Systemstart)" }
    @{ ID=6006;  Cat="Information"; Icon="○"; Desc="Ereignisprotokoll-Dienst beendet (Shutdown)" }
    @{ ID=6008;  Cat="Fehler";      Icon="●"; Desc="Vorheriger Start war unerwartet (Absturz/Stromausfall)" }
    @{ ID=6013;  Cat="Information"; Icon="○"; Desc="System-Uptime-Bericht" }
    @{ ID=7000;  Cat="Fehler";      Icon="●"; Desc="Dienst konnte nicht gestartet werden" }
    @{ ID=7001;  Cat="Fehler";      Icon="●"; Desc="Dienst-Abhängigkeit fehlgeschlagen" }
    @{ ID=7034;  Cat="Fehler";      Icon="●"; Desc="Dienst unerwartet beendet" }
    @{ ID=7036;  Cat="Information"; Icon="○"; Desc="Dienst gestartet oder beendet" }
    @{ ID=7040;  Cat="Information"; Icon="◆"; Desc="Starttyp eines Dienstes geändert" }
    @{ ID=7045;  Cat="Information"; Icon="◆"; Desc="Neuer Dienst installiert (wichtig für Security!)" }
    @{ ID=10016; Cat="Warnung";     Icon="▲"; Desc="DCOM: Startberechtigung für Anwendung verweigert" }
    @{ ID=19;    Cat="Information"; Icon="○"; Desc="Windows Update erfolgreich installiert" }
    @{ ID=20;    Cat="Fehler";      Icon="●"; Desc="Windows Update fehlgeschlagen" }
    @{ ID=43;    Cat="Information"; Icon="◆"; Desc="Windows Update: Installation gestartet" }
    @{ ID=44;    Cat="Information"; Icon="◆"; Desc="Windows Update: Download gestartet" }
    @{ ID=1085;  Cat="Fehler";      Icon="●"; Desc="Gruppenrichtlinie konnte nicht angewendet werden" }
    @{ ID=1500;  Cat="Information"; Icon="○"; Desc="Gruppenrichtlinie erfolgreich verarbeitet" }
    @{ ID=1502;  Cat="Information"; Icon="○"; Desc="Gruppenrichtlinie (Computer) angewendet" }
)
foreach ($e in $sysEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log="System"; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── Security-Events ──────────────────────────────────────────
$secEvents = @(
    @{ ID=4624;  Cat="Information"; Icon="○"; Desc="Erfolgreiche Anmeldung (Logon Type beachten)" }
    @{ ID=4625;  Cat="Warnung";     Icon="▲"; Desc="Fehlgeschlagene Anmeldung" }
    @{ ID=4634;  Cat="Information"; Icon="○"; Desc="Abmeldung eines Kontos" }
    @{ ID=4647;  Cat="Information"; Icon="○"; Desc="Benutzer hat Abmeldung initiiert" }
    @{ ID=4648;  Cat="Warnung";     Icon="▲"; Desc="Anmeldung mit expliziten Credentials (RunAs)" }
    @{ ID=4656;  Cat="Überwachung"; Icon="◆"; Desc="Zugriff auf Objekt angefordert" }
    @{ ID=4663;  Cat="Überwachung"; Icon="◆"; Desc="Zugriff auf Objekt durchgeführt" }
    @{ ID=4670;  Cat="Warnung";     Icon="▲"; Desc="Berechtigungen auf Objekt geändert" }
    @{ ID=4672;  Cat="Information"; Icon="◆"; Desc="Spezielle Berechtigungen vergeben (Admin-Login)" }
    @{ ID=4688;  Cat="Überwachung"; Icon="◆"; Desc="Neuer Prozess erstellt" }
    @{ ID=4689;  Cat="Überwachung"; Icon="◆"; Desc="Prozess beendet" }
    @{ ID=4697;  Cat="Warnung";     Icon="▲"; Desc="Dienst installiert (Security-Log)" }
    @{ ID=4698;  Cat="Warnung";     Icon="▲"; Desc="Geplante Aufgabe erstellt" }
    @{ ID=4699;  Cat="Warnung";     Icon="▲"; Desc="Geplante Aufgabe gelöscht" }
    @{ ID=4700;  Cat="Warnung";     Icon="▲"; Desc="Geplante Aufgabe aktiviert" }
    @{ ID=4701;  Cat="Warnung";     Icon="▲"; Desc="Geplante Aufgabe deaktiviert" }
    @{ ID=4702;  Cat="Warnung";     Icon="▲"; Desc="Geplante Aufgabe aktualisiert" }
    @{ ID=4720;  Cat="Warnung";     Icon="▲"; Desc="Neues Benutzerkonto erstellt" }
    @{ ID=4722;  Cat="Information"; Icon="○"; Desc="Benutzerkonto aktiviert" }
    @{ ID=4723;  Cat="Warnung";     Icon="▲"; Desc="Passwortänderung versucht" }
    @{ ID=4724;  Cat="Warnung";     Icon="▲"; Desc="Passwort zurückgesetzt" }
    @{ ID=4725;  Cat="Warnung";     Icon="▲"; Desc="Benutzerkonto deaktiviert" }
    @{ ID=4726;  Cat="Kritisch";    Icon="●"; Desc="Benutzerkonto gelöscht" }
    @{ ID=4728;  Cat="Warnung";     Icon="▲"; Desc="Mitglied zu globaler Sicherheitsgruppe hinzugefügt" }
    @{ ID=4729;  Cat="Warnung";     Icon="▲"; Desc="Mitglied aus globaler Sicherheitsgruppe entfernt" }
    @{ ID=4732;  Cat="Warnung";     Icon="▲"; Desc="Mitglied zur lokalen Gruppe hinzugefügt" }
    @{ ID=4733;  Cat="Warnung";     Icon="▲"; Desc="Mitglied aus lokaler Gruppe entfernt" }
    @{ ID=4738;  Cat="Warnung";     Icon="▲"; Desc="Benutzerkonto geändert" }
    @{ ID=4740;  Cat="Kritisch";    Icon="●"; Desc="Benutzerkonto gesperrt (Lockout)" }
    @{ ID=4756;  Cat="Warnung";     Icon="▲"; Desc="Mitglied zu universeller Gruppe hinzugefügt" }
    @{ ID=4767;  Cat="Information"; Icon="◆"; Desc="Benutzerkonto entsperrt" }
    @{ ID=4768;  Cat="Überwachung"; Icon="◆"; Desc="Kerberos-TGT angefordert" }
    @{ ID=4769;  Cat="Überwachung"; Icon="◆"; Desc="Kerberos-Service-Ticket angefordert" }
    @{ ID=4771;  Cat="Warnung";     Icon="▲"; Desc="Kerberos-Vorauthentifizierung fehlgeschlagen" }
    @{ ID=4776;  Cat="Warnung";     Icon="▲"; Desc="NTLM-Anmeldung fehlgeschlagen" }
    @{ ID=4797;  Cat="Information"; Icon="◆"; Desc="Abfrage auf leeres Passwort" }
    @{ ID=4946;  Cat="Warnung";     Icon="▲"; Desc="Firewall-Ausnahmeregel hinzugefügt" }
    @{ ID=4947;  Cat="Warnung";     Icon="▲"; Desc="Firewall-Ausnahmeregel geändert" }
    @{ ID=4948;  Cat="Warnung";     Icon="▲"; Desc="Firewall-Ausnahmeregel gelöscht" }
    @{ ID=5136;  Cat="Überwachung"; Icon="◆"; Desc="Verzeichnisdienst-Objekt geändert" }
    @{ ID=5140;  Cat="Überwachung"; Icon="◆"; Desc="Netzwerkfreigabe aufgerufen" }
    @{ ID=5141;  Cat="Warnung";     Icon="▲"; Desc="Verzeichnisdienst-Objekt gelöscht" }
    @{ ID=5156;  Cat="Überwachung"; Icon="◆"; Desc="Windows-Firewall hat Verbindung erlaubt" }
    @{ ID=5157;  Cat="Warnung";     Icon="▲"; Desc="Windows-Firewall hat Verbindung blockiert" }
)
foreach ($e in $secEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log="Security"; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── Application-Events ───────────────────────────────────────
$appEvents = @(
    @{ ID=1000;  Cat="Fehler";      Icon="●"; Desc="Anwendungsabsturz (faulting application)" }
    @{ ID=1001;  Cat="Information"; Icon="◆"; Desc="Windows Error Reporting: Absturzbericht" }
    @{ ID=1002;  Cat="Fehler";      Icon="●"; Desc="Anwendung reagiert nicht (Hang)" }
    @{ ID=1026;  Cat="Fehler";      Icon="●"; Desc=".NET Runtime-Fehler (unbehandelte Ausnahme)" }
    @{ ID=11707; Cat="Information"; Icon="○"; Desc="Anwendung erfolgreich installiert (MSI)" }
    @{ ID=11708; Cat="Fehler";      Icon="●"; Desc="Installation fehlgeschlagen (MSI)" }
    @{ ID=11724; Cat="Information"; Icon="▲"; Desc="Anwendung deinstalliert (MSI)" }
)
foreach ($e in $appEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log="Application"; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── PowerShell-Events ────────────────────────────────────────
$psLog = "Microsoft-Windows-PowerShell/Operational"
$psEvents = @(
    @{ ID=4100; Cat="Fehler";      Icon="●"; Desc="PowerShell-Fehler während Ausführung" }
    @{ ID=4103; Cat="Überwachung"; Icon="◆"; Desc="PowerShell Module Logging (Pipeline-Ausführung)" }
    @{ ID=4104; Cat="Überwachung"; Icon="◆"; Desc="PowerShell Script Block Logging (ausgeführter Code)" }
    @{ ID=4105; Cat="Information"; Icon="○"; Desc="Script Block Invocation Start" }
    @{ ID=4106; Cat="Information"; Icon="○"; Desc="Script Block Invocation Stop" }
    @{ ID=40961;Cat="Warnung";     Icon="▲"; Desc="PowerShell Console-Start mit Problem" }
    @{ ID=40962;Cat="Warnung";     Icon="▲"; Desc="PowerShell Console-Start ohne Profil" }
    @{ ID=53504;Cat="Information"; Icon="◆"; Desc="PowerShell Session-Konfiguration verwendet" }
)
foreach ($e in $psEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$psLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── Windows Defender ─────────────────────────────────────────
$defLog = "Microsoft-Windows-Windows Defender/Operational"
$defEvents = @(
    @{ ID=1006; Cat="Kritisch";    Icon="●"; Desc="Defender: Schadsoftware erkannt" }
    @{ ID=1007; Cat="Warnung";     Icon="▲"; Desc="Defender: Aktion gegen Schadsoftware durchgeführt" }
    @{ ID=1008; Cat="Fehler";      Icon="●"; Desc="Defender: Aktion gegen Schadsoftware fehlgeschlagen" }
    @{ ID=1009; Cat="Information"; Icon="◆"; Desc="Defender: Element aus Quarantäne wiederhergestellt" }
    @{ ID=1015; Cat="Warnung";     Icon="▲"; Desc="Defender: Verdächtiges Verhalten erkannt" }
    @{ ID=1116; Cat="Kritisch";    Icon="●"; Desc="Defender: Malware-Erkennung (Echtzeitschutz)" }
    @{ ID=1117; Cat="Warnung";     Icon="▲"; Desc="Defender: Aktion auf Malware ausgeführt" }
    @{ ID=2000; Cat="Information"; Icon="○"; Desc="Defender: Signatur aktualisiert" }
    @{ ID=2001; Cat="Fehler";      Icon="●"; Desc="Defender: Signatur-Update fehlgeschlagen" }
    @{ ID=5001; Cat="Kritisch";    Icon="●"; Desc="Defender: Echtzeitschutz deaktiviert" }
    @{ ID=5007; Cat="Warnung";     Icon="▲"; Desc="Defender: Konfiguration geändert" }
    @{ ID=5010; Cat="Warnung";     Icon="▲"; Desc="Defender: Scannen deaktiviert" }
)
foreach ($e in $defEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$defLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── WLAN / Netzwerk ──────────────────────────────────────────
$wlanLog = "Microsoft-Windows-WLAN-AutoConfig/Operational"
$wlanEvents = @(
    @{ ID=8001;  Cat="Information"; Icon="○"; Desc="WLAN: Erfolgreich verbunden" }
    @{ ID=8002;  Cat="Fehler";      Icon="●"; Desc="WLAN: Verbindung fehlgeschlagen" }
    @{ ID=8003;  Cat="Information"; Icon="◆"; Desc="WLAN: Verbindung getrennt" }
    @{ ID=11001; Cat="Information"; Icon="○"; Desc="WLAN: Authentifizierung erfolgreich" }
    @{ ID=11004; Cat="Warnung";     Icon="▲"; Desc="WLAN: Authentifizierung fehlgeschlagen" }
    @{ ID=11006; Cat="Warnung";     Icon="▲"; Desc="WLAN: Explizite EAP-Authentifizierung fehlgeschlagen" }
)
foreach ($e in $wlanEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$wlanLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── Task Scheduler ───────────────────────────────────────────
$tsLog = "Microsoft-Windows-TaskScheduler/Operational"
$tsEvents = @(
    @{ ID=106; Cat="Information"; Icon="◆"; Desc="Geplante Aufgabe registriert" }
    @{ ID=129; Cat="Information"; Icon="○"; Desc="Task: Prozess für Aufgabe gestartet" }
    @{ ID=140; Cat="Information"; Icon="◆"; Desc="Geplante Aufgabe aktualisiert" }
    @{ ID=141; Cat="Warnung";     Icon="▲"; Desc="Geplante Aufgabe gelöscht" }
    @{ ID=200; Cat="Information"; Icon="○"; Desc="Task: Aktion gestartet" }
    @{ ID=201; Cat="Information"; Icon="○"; Desc="Task: Aktion abgeschlossen" }
    @{ ID=203; Cat="Fehler";      Icon="●"; Desc="Task: Aktion fehlgeschlagen" }
)
foreach ($e in $tsEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$tsLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── Remote Desktop / Terminaldienste ─────────────────────────
$rdpLsmLog = "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational"
$rdpLsmEvents = @(
    @{ ID=21; Cat="Information"; Icon="○"; Desc="RDP: Sitzungsanmeldung erfolgreich" }
    @{ ID=22; Cat="Information"; Icon="◆"; Desc="RDP: Shell-Start-Benachrichtigung" }
    @{ ID=23; Cat="Information"; Icon="◆"; Desc="RDP: Sitzungsabmeldung" }
    @{ ID=24; Cat="Information"; Icon="◆"; Desc="RDP: Sitzung getrennt (disconnect)" }
    @{ ID=25; Cat="Information"; Icon="○"; Desc="RDP: Sitzung wieder verbunden (reconnect)" }
    @{ ID=39; Cat="Warnung";     Icon="▲"; Desc="RDP: Sitzung wurde von anderer Sitzung getrennt" }
    @{ ID=40; Cat="Information"; Icon="◆"; Desc="RDP: Sitzungsverbindung getrennt" }
)
foreach ($e in $rdpLsmEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$rdpLsmLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

$rdpRcmLog = "Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational"
$rdpRcmEvents = @(
    @{ ID=1149; Cat="Information"; Icon="◆"; Desc="RDP: Benutzerauthentifizierung erfolgreich (Vor-Logon)" }
    @{ ID=1150; Cat="Warnung";     Icon="▲"; Desc="RDP: Authentifizierung mit Fehler abgeschlossen" }
    @{ ID=1158; Cat="Information"; Icon="○"; Desc="RDP: Remote-Verbindung aufgebaut" }
)
foreach ($e in $rdpRcmEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$rdpRcmLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── Druckerdienst ────────────────────────────────────────────
$prtLog = "Microsoft-Windows-PrintService/Operational"
$prtEvents = @(
    @{ ID=307; Cat="Information"; Icon="○"; Desc="Druckauftrag gedruckt (Doc, Benutzer, Drucker)" }
    @{ ID=805; Cat="Information"; Icon="◆"; Desc="Druckauftrag gelöscht" }
    @{ ID=842; Cat="Warnung";     Icon="▲"; Desc="Druckauftrag fehlgeschlagen" }
)
foreach ($e in $prtEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$prtLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── BitLocker ────────────────────────────────────────────────
$blLog = "Microsoft-Windows-BitLocker/BitLocker Management"
$blEvents = @(
    @{ ID=845; Cat="Information"; Icon="○"; Desc="BitLocker: Verschlüsselung aktiviert" }
    @{ ID=846; Cat="Information"; Icon="◆"; Desc="BitLocker: Verschlüsselung deaktiviert" }
    @{ ID=24620; Cat="Warnung";   Icon="▲"; Desc="BitLocker: Recovery-Modus ausgelöst" }
)
foreach ($e in $blEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$blLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ─── SMB Client (Netzwerkfreigaben) ───────────────────────────
$smbLog = "Microsoft-Windows-SmbClient/Connectivity"
$smbEvents = @(
    @{ ID=30803; Cat="Warnung"; Icon="▲"; Desc="SMB: Verbindung zum Server unterbrochen" }
    @{ ID=30806; Cat="Warnung"; Icon="▲"; Desc="SMB: Netzwerk-Signatur-Fehler" }
)
foreach ($e in $smbEvents) { $eventCatalog.Add([PSCustomObject]@{ ID=$e.ID; Log=$smbLog; Category=$e.Cat; Icon=$e.Icon; Desc=$e.Desc }) }

# ════════════════════════════════════════════════════════════
#  COMPUTER-SCAN: Alle vorhandenen Event-IDs erkennen
# ════════════════════════════════════════════════════════════
# Liste aller bereits im Katalog vorhandenen Logs (für die Bereichs-Auswahl)
$script:discoveredLogs = [System.Collections.Generic.List[string]]::new()

function New-ProgressForm {
    param([string]$ComputerName)

    $pf = New-Object System.Windows.Forms.Form
    $pf.Text            = "Event-Logs werden gescannt"
    $pf.Size            = New-Object System.Drawing.Size(540, 220)
    $pf.StartPosition   = "CenterScreen"
    $pf.FormBorderStyle = "FixedDialog"
    $pf.MaximizeBox     = $false
    $pf.MinimizeBox     = $false
    $pf.BackColor       = $clrBg
    $pf.Font            = $fontNormal
    $pf.ControlBox      = $false

    $pnlHead = New-Object System.Windows.Forms.Panel
    $pnlHead.Dock      = "Top"
    $pnlHead.Height    = 50
    $pnlHead.BackColor = $clrAccent
    $pf.Controls.Add($pnlHead)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Text      = "  Scanne Computer: $ComputerName"
    $lblHead.Dock      = "Fill"
    $lblHead.Font      = $fontTitle
    $lblHead.ForeColor = [System.Drawing.Color]::White
    $lblHead.TextAlign = "MiddleLeft"
    $pnlHead.Controls.Add($lblHead)

    $lblMain = New-Object System.Windows.Forms.Label
    $lblMain.Location  = New-Object System.Drawing.Point(20, 65)
    $lblMain.Size      = New-Object System.Drawing.Size(490, 20)
    $lblMain.Text      = "Verbindung wird hergestellt..."
    $lblMain.Font      = $fontBold
    $pf.Controls.Add($lblMain)

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Location  = New-Object System.Drawing.Point(20, 88)
    $lblSub.Size      = New-Object System.Drawing.Size(490, 18)
    $lblSub.Text      = ""
    $lblSub.Font      = $fontSmall
    $lblSub.ForeColor = $clrMuted
    $pf.Controls.Add($lblSub)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(20, 112)
    $bar.Size     = New-Object System.Drawing.Size(490, 16)
    $bar.Style    = "Continuous"
    $bar.Minimum  = 0
    $bar.Maximum  = 100
    $pf.Controls.Add($bar)

    $btnSkip = New-StyledButton "Überspringen" 410 142 100 28 $false
    $pf.Controls.Add($btnSkip)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location  = New-Object System.Drawing.Point(20, 145)
    $lblHint.Size      = New-Object System.Drawing.Size(380, 32)
    $lblHint.Text      = "Es werden alle aktiven Event-Logs durchsucht.`nDas kann 10-60 Sekunden dauern."
    $lblHint.Font      = $fontSmall
    $lblHint.ForeColor = $clrMuted
    $pf.Controls.Add($lblHint)

    $script:scanCancelled = $false
    $btnSkip.Add_Click({ $script:scanCancelled = $true })

    return [PSCustomObject]@{
        Form    = $pf
        LblMain = $lblMain
        LblSub  = $lblSub
        Bar     = $bar
    }
}

function Invoke-ComputerEventScan {
    param(
        [string]$Computer,
        $ProgressUI,
        [int]$MaxPerLog  = 15,    # Klein halten: nur ID-Erkennung, kein Content nötig
        [int]$MaxLogs    = 120,   # Top-N aktivste Logs – der Rest bringt kaum neue IDs
        [System.Management.Automation.PSCredential]$Credential = $null
    )

    $found = [System.Collections.Generic.List[PSObject]]::new()
    $script:scanCancelled = $false

    # Schritt 1: Alle Logs ermitteln (nur Metadaten – sehr schnell)
    $ProgressUI.LblMain.Text = "Logs werden aufgelistet..."
    $ProgressUI.LblSub.Text  = "Get-WinEvent -ListLog *"
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $listParams = @{ ListLog = '*'; ComputerName = $Computer; ErrorAction = 'Stop' }
        if ($Credential) { $listParams.Credential = $Credential }
        # Nur aktive Logs, absteigend nach RecordCount, auf MaxLogs begrenzen
        $allLogs = Get-WinEvent @listParams |
                   Where-Object { $_.RecordCount -gt 0 -and $_.IsEnabled } |
                   Sort-Object RecordCount -Descending |
                   Select-Object -First $MaxLogs
    } catch {
        $ProgressUI.LblMain.Text = "Fehler: $($_.Exception.Message)"
        Start-Sleep -Seconds 2
        return $null
    }

    if (-not $allLogs -or $allLogs.Count -eq 0) { return $found }

    # Schritt 2: Pro Log kleines Sample lesen – nur zur ID-Erkennung
    # WICHTIG: $event.Message NICHT verwenden – löst teure DLL-Formatierung aus!
    $ProgressUI.Bar.Maximum = $allLogs.Count
    $i = 0
    foreach ($log in $allLogs) {
        if ($script:scanCancelled) { break }
        $i++
        $ProgressUI.Bar.Value    = $i
        $ProgressUI.LblMain.Text = "Scanne Log $i / $($allLogs.Count): $($log.LogName)"
        $ProgressUI.LblSub.Text  = "$($log.RecordCount) Einträge · lese $MaxPerLog Events..."
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $sampleParams = @{
                LogName      = $log.LogName
                MaxEvents    = $MaxPerLog
                ComputerName = $Computer
                ErrorAction  = 'Stop'
            }
            if ($Credential) { $sampleParams.Credential = $Credential }
            $sample = Get-WinEvent @sampleParams

            # Nach ID + Provider gruppieren – kein .Message-Zugriff (wäre langsam)
            $grouped = $sample | Group-Object -Property Id
            foreach ($g in $grouped) {
                $first    = $g.Group[0]
                $level    = $first.LevelDisplayName
                $provider = $first.ProviderName

                $icon = switch ($level) {
                    "Kritisch"      { "●" } "Critical"      { "●" }
                    "Fehler"        { "●" } "Error"         { "●" }
                    "Warnung"       { "▲" } "Warning"       { "▲" }
                    "Information"   { "○" } "Informationen" { "○" }
                    default         { "◈" }
                }

                # Beschreibung ohne .Message → kein DLL-Aufruf
                $desc = "[Provider: $provider · $($g.Count)x in Sample · Log: $($log.LogName)]"

                $found.Add([PSCustomObject]@{
                    ID       = [int]$first.Id
                    Log      = $log.LogName
                    Category = "Erkannt"
                    Icon     = $icon
                    Desc     = $desc
                })
            }
        } catch {
            # Log nicht lesbar (Berechtigung o.ä.) – überspringen
        }
    }

    return $found
}

function Invoke-DocumentedIDsScan {
    # Liest Provider-Manifeste, um ALLE dokumentierten Event-IDs pro Log zu erfassen
    # (auch solche, die im aktuellen Sample nicht vorgekommen sind).
    param(
        [string]$Computer,
        $ProgressUI,
        [string[]]$LogNames,
        [System.Management.Automation.PSCredential]$Credential = $null
    )

    $found = [System.Collections.Generic.List[PSObject]]::new()
    if (-not $LogNames -or $LogNames.Count -eq 0) { return $found }

    # 1) Provider sammeln, die in diese Logs schreiben
    $providers = New-Object System.Collections.Generic.HashSet[string]
    foreach ($logName in $LogNames) {
        if ($script:scanCancelled) { return $found }
        try {
            $logInfoParams = @{ ListLog = $logName; ComputerName = $Computer; ErrorAction = 'Stop' }
            if ($Credential) { $logInfoParams.Credential = $Credential }
            $logInfo = Get-WinEvent @logInfoParams
            foreach ($p in $logInfo.ProviderNames) {
                [void]$providers.Add($p)
            }
        } catch {}
    }

    # Lookup: nur Events behalten, deren LogLink in unserer Log-Liste ist
    $logSet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($l in $LogNames) { [void]$logSet.Add($l) }

    # 2) Pro Provider Manifest-Events lesen
    $ProgressUI.Bar.Maximum = [Math]::Max(1, $providers.Count)
    $ProgressUI.Bar.Value   = 0
    $i = 0
    $totalProv = $providers.Count

    foreach ($provName in $providers) {
        if ($script:scanCancelled) { break }
        $i++
        $ProgressUI.Bar.Value    = [Math]::Min($i, $ProgressUI.Bar.Maximum)
        $ProgressUI.LblMain.Text = "Manifest $i / $totalProv : $provName"
        $ProgressUI.LblSub.Text  = "Lese dokumentierte Event-IDs aus Provider-Manifest..."
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $provParams = @{ ListProvider = $provName; ComputerName = $Computer; ErrorAction = 'Stop' }
            if ($Credential) { $provParams.Credential = $Credential }
            $prov = Get-WinEvent @provParams
            foreach ($evt in $prov.Events) {
                # Zu welchem Log gehört dieses Event?
                $targetLog = $null
                if ($evt.LogLink -and $evt.LogLink.LogName) {
                    $targetLog = $evt.LogLink.LogName
                }
                if (-not $targetLog -or -not $logSet.Contains($targetLog)) { continue }

                # Level -> Icon
                $levelName = ""
                if ($evt.Level) { $levelName = "$($evt.Level.DisplayName)" }
                $icon = switch -Wildcard ($levelName) {
                    "*ritisch*"     { "●"; break }
                    "*ritical*"     { "●"; break }
                    "*ehler*"       { "●"; break }
                    "*rror*"        { "●"; break }
                    "*arnung*"      { "▲"; break }
                    "*arning*"      { "▲"; break }
                    "*nformation*"  { "○"; break }
                    default         { "◈" }
                }

                # Beschreibung aus Manifest (Template kann %1, %2 enthalten - das ist ok)
                $desc = ""
                if ($evt.Description) {
                    $desc = ($evt.Description -replace "`r`n|`n", " ").Trim()
                }
                if ($desc.Length -gt 110) { $desc = $desc.Substring(0, 110) + "..." }
                if ([string]::IsNullOrWhiteSpace($desc)) { $desc = "(keine Beschreibung im Manifest)" }

                $found.Add([PSCustomObject]@{
                    ID       = [int]$evt.Id
                    Log      = $targetLog
                    Category = "Dokumentiert"
                    Icon     = $icon
                    Desc     = "$desc  [Provider: $provName]"
                })
            }
        } catch {
            # Provider-Manifest nicht lesbar - ignorieren
        }
    }

    return $found
}

function Merge-DiscoveredEvents {
    param([System.Collections.Generic.List[PSObject]]$Discovered)

    if ($null -eq $Discovered -or $Discovered.Count -eq 0) { return 0 }

    # Existing-Lookup: "ID|Log"
    $existing = @{}
    foreach ($ev in $eventCatalog) {
        $existing["$($ev.ID)|$($ev.Log)"] = $true
    }

    $added = 0
    $logsSeen = New-Object System.Collections.Generic.HashSet[string]
    foreach ($ev in $Discovered) {
        $key = "$($ev.ID)|$($ev.Log)"
        if (-not $existing.ContainsKey($key)) {
            $eventCatalog.Add($ev)
            $existing[$key] = $true
            $added++
        }
        [void]$logsSeen.Add($ev.Log)
    }

    # Erkannte Logs für Bereichsfilter merken
    $script:discoveredLogs.Clear()
    foreach ($l in ($logsSeen | Sort-Object)) { $script:discoveredLogs.Add($l) }

    return $added
}

# ════════════════════════════════════════════════════════════
#  STARTUP-SCAN wird weiter unten nach Definition der Helper-Funktionen ausgefuehrt
# ════════════════════════════════════════════════════════════

# ── Hilfsfunktionen ──────────────────────────────────────────
function New-StyledButton($text, $x, $y, $w=120, $h=32, $primary=$true) {
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text      = $text
    $btn.Location  = New-Object System.Drawing.Point($x, $y)
    $btn.Size      = New-Object System.Drawing.Size($w, $h)
    $btn.Font      = $fontBold
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    if ($primary) {
        $btn.BackColor  = $clrAccent
        $btn.ForeColor  = [System.Drawing.Color]::White
        $btn.Add_MouseEnter({ $this.BackColor = $clrAccentHov })
        $btn.Add_MouseLeave({ $this.BackColor = $clrAccent })
    } else {
        $btn.BackColor  = $clrPanel
        $btn.ForeColor  = $clrText
        $btn.FlatAppearance.BorderSize = 1
        $btn.FlatAppearance.BorderColor = $clrBorder
    }
    return $btn
}

function New-Label($text, $x, $y, $w=140, $h=20, $bold=$false, $muted=$false) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $text
    $lbl.Location  = New-Object System.Drawing.Point($x, $y)
    $lbl.Size      = New-Object System.Drawing.Size($w, $h)
    $lbl.Font      = if ($bold) { $fontBold } else { $fontNormal }
    $lbl.ForeColor = if ($muted) { $clrMuted } else { $clrText }
    return $lbl
}

function New-SectionPanel($x, $y, $w, $h, $title) {
    $pnl = New-Object System.Windows.Forms.Panel
    $pnl.Location  = New-Object System.Drawing.Point($x, $y)
    $pnl.Size      = New-Object System.Drawing.Size($w, $h)
    $pnl.BackColor = $clrPanel
    $pnl.BorderStyle = "None"

    # Zeichne Rahmen mit Paint-Event
    $pnl.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen($clrBorder, 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width - 1, $s.Height - 1)
        $pen.Dispose()
    })

    if ($title) {
        $lbl = New-Label $title 10 8 ($w-20) 20 $true
        $lbl.ForeColor = $clrAccent
        $pnl.Controls.Add($lbl)
    }
    return $pnl
}

# ════════════════════════════════════════════════════════════
#  STARTUP-SCAN ausführen (alle Helper sind nun definiert)
# ════════════════════════════════════════════════════════════
$script:startupScanResult = $null
$startupComputer = $env:COMPUTERNAME
try {
    $progUI = New-ProgressForm -ComputerName $startupComputer
    $progUI.Form.Show()
    $progUI.Form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    # Phase 1: Schnell-Scan – kleine Samples, kein Message-Lookup, kein Manifest-Scan
    $discovered = Invoke-ComputerEventScan -Computer $startupComputer -ProgressUI $progUI `
                      -MaxPerLog 15 -MaxLogs 120
    $addedCount = Merge-DiscoveredEvents -Discovered $discovered
    $logsCount  = $script:discoveredLogs.Count

    $progUI.Form.Close()
    $progUI.Form.Dispose()

    if ($script:scanCancelled) {
        $script:startupScanResult = "Scan vom Benutzer abgebrochen."
    } else {
        $script:startupScanResult = "Schnell-Scan: $addedCount IDs aus $logsCount Logs erkannt. Für Manifest-Scan: 'Scan + Manifest' nutzen."
    }
} catch {
    if ($progUI -and $progUI.Form) {
        try { $progUI.Form.Close(); $progUI.Form.Dispose() } catch {}
    }
    $script:startupScanResult = "Scan fehlgeschlagen: $($_.Exception.Message)"
}

# ════════════════════════════════════════════════════════════
#  FORMULAR 1 – Auswahl
# ════════════════════════════════════════════════════════════
$formMain = New-Object System.Windows.Forms.Form
$formMain.Text            = "Windows Event Viewer – Abfrage-Tool"
$formMain.Size            = New-Object System.Drawing.Size(920, 845)
$formMain.StartPosition   = "CenterScreen"
$formMain.BackColor       = $clrBg
$formMain.FormBorderStyle = "FixedSingle"
$formMain.MaximizeBox     = $false
$formMain.Font            = $fontNormal

# ── Titel-Leiste ─────────────────────────────────────────────
$pnlTitle = New-Object System.Windows.Forms.Panel
$pnlTitle.Dock      = "Top"
$pnlTitle.Height    = 60
$pnlTitle.BackColor = $clrAccent
$formMain.Controls.Add($pnlTitle)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "  Windows Event Abfrage-Tool"
$lblTitle.Dock      = "Fill"
$lblTitle.Font      = $fontTitle
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.TextAlign = "MiddleLeft"
$pnlTitle.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text      = "   Mehrfachauswahl möglich – Strg+Klick oder Checkboxen nutzen"
$lblSubtitle.Dock      = "Bottom"
$lblSubtitle.Height    = 18
$lblSubtitle.Font      = $fontSmall
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
$lblSubtitle.BackColor = $clrAccent
$pnlTitle.Controls.Add($lblSubtitle)

# ── Sektion: Filter-Optionen ──────────────────────────────────
$pnlOptions = New-SectionPanel 15 75 885 160 "Abfrage-Optionen"
$formMain.Controls.Add($pnlOptions)

# Zeitraum
$pnlOptions.Controls.Add((New-Label "Zeitraum:" 10 32 80 20))
$cbZeit = New-Object System.Windows.Forms.ComboBox
$cbZeit.Location     = New-Object System.Drawing.Point(95, 30)
$cbZeit.Size         = New-Object System.Drawing.Size(150, 22)
$cbZeit.DropDownStyle = "DropDownList"
$cbZeit.Font         = $fontNormal
@("Letzte 1 Stunde","Letzte 6 Stunden","Letzte 24 Stunden","Letzte 7 Tage","Letzte 30 Tage","Alles") |
    ForEach-Object { $cbZeit.Items.Add($_) | Out-Null }
$cbZeit.SelectedIndex = 2
$pnlOptions.Controls.Add($cbZeit)

# Max. Einträge
$pnlOptions.Controls.Add((New-Label "Max. Einträge:" 255 32 95 20))
$cbMax = New-Object System.Windows.Forms.ComboBox
$cbMax.Location      = New-Object System.Drawing.Point(355, 30)
$cbMax.Size          = New-Object System.Drawing.Size(70, 22)
$cbMax.DropDownStyle = "DropDownList"
$cbMax.Font          = $fontNormal
@(25, 50, 100, 250, 500) | ForEach-Object { $cbMax.Items.Add($_) | Out-Null }
$cbMax.SelectedIndex = 1
$pnlOptions.Controls.Add($cbMax)

# Computer
$pnlOptions.Controls.Add((New-Label "Computer:" 435 32 70 20))
$txtComputer = New-Object System.Windows.Forms.TextBox
$txtComputer.Location  = New-Object System.Drawing.Point(508, 30)
$txtComputer.Size      = New-Object System.Drawing.Size(198, 22)
$txtComputer.Font      = $fontNormal
$txtComputer.Text      = $env:COMPUTERNAME
$txtComputer.ForeColor = $clrText
$pnlOptions.Controls.Add($txtComputer)

# Re-Scan Button + Manifest-Option
$btnRescan = New-StyledButton "Scan" 714 28 80 26 $false
$pnlOptions.Controls.Add($btnRescan)

$chkManifest = New-Object System.Windows.Forms.CheckBox
$chkManifest.Location = New-Object System.Drawing.Point(800, 30)
$chkManifest.Size     = New-Object System.Drawing.Size(18, 18)
$pnlOptions.Controls.Add($chkManifest)
$lblManifest = New-Label "Manifest" 820 32 52 16 $false $true
$lblManifest.Font = $fontSmall
$pnlOptions.Controls.Add($lblManifest)

# ── Credentials (Zeile 2) ─────────────────────────────────────
$pnlOptions.Controls.Add((New-Label "Domain:" 10 70 55 20))
$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Location  = New-Object System.Drawing.Point(68, 68)
$txtDomain.Size      = New-Object System.Drawing.Size(120, 22)
$txtDomain.Font      = $fontNormal
$pnlOptions.Controls.Add($txtDomain)

$pnlOptions.Controls.Add((New-Label "Benutzer:" 205 70 65 20))
$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location  = New-Object System.Drawing.Point(273, 68)
$txtUser.Size      = New-Object System.Drawing.Size(155, 22)
$txtUser.Font      = $fontNormal
$pnlOptions.Controls.Add($txtUser)

$pnlOptions.Controls.Add((New-Label "Passwort:" 440 70 65 20))
$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location     = New-Object System.Drawing.Point(508, 68)
$txtPass.Size         = New-Object System.Drawing.Size(155, 22)
$txtPass.Font         = $fontNormal
$txtPass.PasswordChar = '*'
$pnlOptions.Controls.Add($txtPass)

$lblCredHint = New-Label "Leer = aktuelle Windows-Anmeldung" 671 71 205 18 $false $true
$lblCredHint.Font = $fontSmall
$pnlOptions.Controls.Add($lblCredHint)

# ── Zeile 3: Auswahl-Profile ──────────────────────────────────
$pnlOptions.Controls.Add((New-Label "Profil:" 10 112 65 20 $true))
$cbProfile = New-Object System.Windows.Forms.ComboBox
$cbProfile.Location      = New-Object System.Drawing.Point(78, 110)
$cbProfile.Size          = New-Object System.Drawing.Size(145, 22)
$cbProfile.Font          = $fontNormal
$cbProfile.DropDownStyle = "DropDown"
$pnlOptions.Controls.Add($cbProfile)

$btnLoadProfile = New-StyledButton "Laden"     230 108 85 26 $false
$btnSaveProfile = New-StyledButton "Speichern" 322 108 100 26 $false
$pnlOptions.Controls.Add($btnLoadProfile)
$pnlOptions.Controls.Add($btnSaveProfile)

$btnLoadProfile.Add_Click({ Load-Profile $cbProfile.Text })
$btnSaveProfile.Add_Click({ Save-Profile $cbProfile.Text })

# ── Zeile 3: Live-Modus ───────────────────────────────────────
$pnlOptions.Controls.Add((New-Label "Live-Modus:" 432 112 80 20 $true))
$chkLive = New-Object System.Windows.Forms.CheckBox
$chkLive.Location = New-Object System.Drawing.Point(516, 111)
$chkLive.Size     = New-Object System.Drawing.Size(18, 18)
$pnlOptions.Controls.Add($chkLive)

$cbLiveInterval = New-Object System.Windows.Forms.ComboBox
$cbLiveInterval.Location      = New-Object System.Drawing.Point(538, 110)
$cbLiveInterval.Size          = New-Object System.Drawing.Size(105, 22)
$cbLiveInterval.DropDownStyle = "DropDownList"
$cbLiveInterval.Font          = $fontNormal
@("30 Sek","1 Minute","5 Minuten","10 Minuten") | ForEach-Object { $cbLiveInterval.Items.Add($_) | Out-Null }
$cbLiveInterval.SelectedIndex = 1
$pnlOptions.Controls.Add($cbLiveInterval)

$lblLiveHint = New-Label "Autom. Aktualisierung der Ergebnisse" 650 113 220 18 $false $true
$lblLiveHint.Font = $fontSmall
$pnlOptions.Controls.Add($lblLiveHint)

# Profilliste initialisieren
Refresh-ProfileList

# ── Tooltips für alle Steuerelemente ─────────────────────────
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 9000
$toolTip.InitialDelay = 450
$toolTip.ReshowDelay  = 200
$toolTip.ShowAlways   = $true

$toolTip.SetToolTip($cbZeit,
    "Zeitraum der abzufragenden Windows-Ereignisse.`n" +
    "Kleiner Zeitraum = schnellere Abfrage.`n`n" +
    "Empfehlung: 'Letzte 24 Stunden' für schnellen Überblick.")

$toolTip.SetToolTip($cbMax,
    "Maximale Anzahl Einträge pro Log-Abfrage.`n" +
    "Kleinerer Wert = schnellere Abfrage, weniger Treffer.`n`n" +
    "Hinweis: Bei Mehrfach-Auswahl wird dieser Wert je Event-Gruppe angewendet.")

$toolTip.SetToolTip($txtComputer,
    "Hostname oder IP-Adresse des Ziel-Computers.`n" +
    "Mehrere Computer kommagetrennt eingeben: PC1, PC2, SRV01`n`n" +
    "Leer lassen = lokaler Computer wird abgefragt.")

$toolTip.SetToolTip($btnRescan,
    "Schnell-Scan des Computers: erkennt aktiv vorhandene Event-IDs`n" +
    "in den wichtigsten Logs und ergänzt die 'Erkannt'-Liste.`n`n" +
    "Manifest-Checkbox aktivieren für tieferen Scan.")

$toolTip.SetToolTip($chkManifest,
    "Manifest-Scan aktivieren: liest zusätzlich alle dokumentierten`n" +
    "Event-IDs aus Provider-Manifesten (langsamer, umfassender).")

$toolTip.SetToolTip($lblManifest,
    "Manifest-Scan aktivieren: liest zusätzlich alle dokumentierten`n" +
    "Event-IDs aus Provider-Manifesten (langsamer, umfassender).")

$toolTip.SetToolTip($txtDomain,
    "Active-Directory-Domäne für den Remote-Zugriff.`n" +
    "Beispiel: FIRMA oder firma.local`n`n" +
    "Leer lassen = lokales Konto wird verwendet.")

$toolTip.SetToolTip($txtUser,
    "Benutzername für den Remote-Zugriff auf den Ziel-Computer.`n`n" +
    "Leer lassen = aktuelle Windows-Anmeldung wird verwendet.")

$toolTip.SetToolTip($txtPass,
    "Passwort für den Remote-Zugriff (wird nicht gespeichert).`n`n" +
    "Leer lassen = aktuelle Windows-Anmeldung wird verwendet.")

$toolTip.SetToolTip($cbProfile,
    "Name des Abfrage-Profils.`n" +
    "Profile speichern die Event-Auswahl, den Zeitraum und das Max-Limit.`n`n" +
    "Namen eingeben und 'Speichern' klicken, um ein neues Profil anzulegen.")

$toolTip.SetToolTip($btnLoadProfile,
    "Gespeichertes Profil laden:`n" +
    "Stellt die gespeicherte Event-Auswahl, Zeitraum und Max-Limit wieder her.")

$toolTip.SetToolTip($btnSaveProfile,
    "Aktuelle Event-Auswahl als Profil speichern.`n" +
    "Profile werden lokal als JSON-Dateien gespeichert.")

$toolTip.SetToolTip($chkLive,
    "Live-Modus aktivieren: Die Ergebnisliste wird automatisch`n" +
    "im eingestellten Zeitintervall aktualisiert.`n`n" +
    "Hinweis: Erhöht die Last auf dem Ziel-Computer.")

$toolTip.SetToolTip($cbLiveInterval,
    "Zeitintervall für die automatische Aktualisierung im Live-Modus.`n" +
    "Kürzere Intervalle = aktueller, aber mehr Netzwerkverkehr.")

function Get-FormCredential {
    $user = $txtUser.Text.Trim()
    $pass = $txtPass.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrWhiteSpace($pass)) { return $null }
    $domain = $txtDomain.Text.Trim()
    $loginName = if ($domain -ne "") { "$domain\$user" } else { $user }
    $secPass = ConvertTo-SecureString $pass -AsPlainText -Force
    return New-Object System.Management.Automation.PSCredential($loginName, $secPass)
}

# ── Profil-Funktionen ─────────────────────────────────────────
function Save-Profile {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        [System.Windows.Forms.MessageBox]::Show("Bitte einen Profilnamen eingeben.",
            "Kein Name", "OK", "Warning") | Out-Null
        return
    }
    $events = [System.Collections.Generic.List[hashtable]]::new()
    for ($i = 0; $i -lt $clbEvents.Items.Count; $i++) {
        if ($clbEvents.GetItemChecked($i) -and $i -lt $script:visibleEvents.Count) {
            $ev = $script:visibleEvents[$i]
            $events.Add(@{ ID = $ev.ID; Log = $ev.Log })
        }
    }
    if ($events.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Keine Events ausgewählt – Profil wird nicht gespeichert.",
            "Leer", "OK", "Warning") | Out-Null
        return
    }
    $obj = [ordered]@{
        Name     = $Name
        Zeitraum = $cbZeit.SelectedIndex
        MaxCount = $cbMax.SelectedIndex
        Events   = @($events)
    }
    $path = Join-Path $script:profileDir "$Name.json"
    $obj | ConvertTo-Json -Depth 5 | Set-Content $path -Encoding UTF8
    Refresh-ProfileList
    $cbProfile.Text = $Name
    [System.Windows.Forms.MessageBox]::Show("Profil '$Name' gespeichert ($($events.Count) Events).",
        "Gespeichert", "OK", "Information") | Out-Null
}

function Load-Profile {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return }
    $path = Join-Path $script:profileDir "$Name.json"
    if (-not (Test-Path $path)) {
        [System.Windows.Forms.MessageBox]::Show("Profil '$Name' nicht gefunden.", "Fehler", "OK", "Warning") | Out-Null
        return
    }
    $prof = Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json

    # Filter zurücksetzen
    $cbZeit.SelectedIndex = [int]$prof.Zeitraum
    $cbMax.SelectedIndex  = [int]$prof.MaxCount

    # Alle Bereiche anzeigen damit alle Profil-IDs sichtbar sind
    $cbKat.SelectedIndex = 0
    $cbLog.SelectedIndex = 0

    # Events-Liste neu befüllen
    Update-EventList

    # Profil-Events ankreuzen
    $profSet = @{}
    foreach ($e in $prof.Events) { $profSet["$($e.ID)|$($e.Log)"] = $true }

    for ($i = 0; $i -lt $script:visibleEvents.Count; $i++) {
        $ev  = $script:visibleEvents[$i]
        $key = "$($ev.ID)|$($ev.Log)"
        $clbEvents.SetItemChecked($i, $profSet.ContainsKey($key))
    }

    $hitCount = ($profSet.Keys | Where-Object { $profSet[$_] }).Count
    $lblStatus.ForeColor = $clrSuccess
    $lblStatus.Text = "✔ Profil '$Name' geladen – $($hitCount) Events ausgewählt."
}

function Refresh-ProfileList {
    $current = $cbProfile.Text
    $cbProfile.Items.Clear()
    Get-ChildItem $script:profileDir -Filter "*.json" -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty BaseName |
        Sort-Object |
        ForEach-Object { $cbProfile.Items.Add($_) | Out-Null }
    if ($cbProfile.Items.Contains($current)) { $cbProfile.Text = $current }
}

# ── Live-Timer (Skript-Scope für FormOut-Zugriff) ─────────────
$script:liveTimer    = New-Object System.Windows.Forms.Timer
$script:liveTimer.Interval = 60000   # Default 1 Minute
$script:liveDgv      = $null         # Referenz auf aktives DataGridView
$script:liveDataRef  = $null         # Referenz auf $sorted der aktuellen Abfrage

# ── Sektion: Event-Auswahl ────────────────────────────────────
$pnlEvents = New-SectionPanel 15 245 885 370 "Events auswählen  (Mehrfachauswahl per Checkbox)"
$formMain.Controls.Add($pnlEvents)

# Kategorie-Filter Dropdown
$pnlEvents.Controls.Add((New-Label "Kategorie:" 10 35 75 20))
$cbKat = New-Object System.Windows.Forms.ComboBox
$cbKat.Location      = New-Object System.Drawing.Point(88, 33)
$cbKat.Size          = New-Object System.Drawing.Size(130, 22)
$cbKat.DropDownStyle = "DropDownList"
$cbKat.Font          = $fontNormal
@("Alle Kategorien","Kritisch","Fehler","Warnung","Information","Überwachung","Eigene","Erkannt","Dokumentiert") |
    ForEach-Object { $cbKat.Items.Add($_) | Out-Null }
$cbKat.SelectedIndex = 0
$pnlEvents.Controls.Add($cbKat)

# Log-Gruppe Filter Dropdown (nur Anzeige-Filter)
$pnlEvents.Controls.Add((New-Label "Bereich:" 235 35 65 20))
$cbLog = New-Object System.Windows.Forms.ComboBox
$cbLog.Location      = New-Object System.Drawing.Point(300, 33)
$cbLog.Size          = New-Object System.Drawing.Size(220, 22)
$cbLog.DropDownStyle = "DropDownList"
$cbLog.Font          = $fontNormal
@(
    "Alle Bereiche",
    ">> Gefunden (alle Scan-Ergebnisse)",
    "★ Empfohlen (kuratiert)",
    "◈ Erkannt auf Computer",
    "= Dokumentiert (Manifest)",
    "System",
    "Security",
    "Application",
    "PowerShell",
    "Windows Defender",
    "WLAN / Netzwerk",
    "Task Scheduler",
    "Remote Desktop",
    "Druckerdienst",
    "BitLocker",
    "SMB Client",
    "Eigene IDs"
) | ForEach-Object { $cbLog.Items.Add($_) | Out-Null }
$cbLog.SelectedIndex = 0
$pnlEvents.Controls.Add($cbLog)

# Alle/Keine Buttons
$btnAll  = New-StyledButton "[+] Alle"  545 30 85 24 $false
$btnNone = New-StyledButton "[-] Keine" 640 30 85 24 $false
$pnlEvents.Controls.Add($btnAll)
$pnlEvents.Controls.Add($btnNone)

# Zähler rechts
$lblCatCount = New-Label "" 740 35 130 20 $false $true
$pnlEvents.Controls.Add($lblCatCount)

# ── Volltextsuche ─────────────────────────────────────────────
$script:searchFilter = ""
$pnlEvents.Controls.Add((New-Label "Suche:" 10 64 48 20))
$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location    = New-Object System.Drawing.Point(60, 62)
$txtSearch.Size        = New-Object System.Drawing.Size(700, 22)
$txtSearch.Font        = $fontNormal
$txtSearch.BackColor   = $clrPanel
$txtSearch.ForeColor   = $clrText
$txtSearch.BorderStyle = "FixedSingle"
$pnlEvents.Controls.Add($txtSearch)
$btnSearchClear = New-StyledButton "X" 766 60 30 24 $false
$pnlEvents.Controls.Add($btnSearchClear)
$toolTip.SetToolTip($txtSearch,
    "Volltextsuche über Event-ID, Beschreibung, Kategorie und Log-Bereich.`n" +
    "Die Liste wird sofort beim Tippen gefiltert.")
$toolTip.SetToolTip($btnSearchClear, "Suchfeld leeren")

# CheckedListBox für Events – monospace für saubere Ausrichtung
$clbEvents = New-Object System.Windows.Forms.CheckedListBox
$clbEvents.Location       = New-Object System.Drawing.Point(10, 90)
$clbEvents.Size           = New-Object System.Drawing.Size(860, 267)
$clbEvents.Font           = New-Object System.Drawing.Font("Consolas", 9)
$clbEvents.BackColor      = $clrPanel
$clbEvents.ForeColor      = $clrText
$clbEvents.BorderStyle    = "None"
$clbEvents.CheckOnClick   = $true
$clbEvents.HorizontalScrollbar = $true
$pnlEvents.Controls.Add($clbEvents)

# Mapping Log-Gruppenname -> tatsächliche Log-Namen (für Filter)
$script:logGroupMap = @{
    "System"           = @("System")
    "Security"         = @("Security")
    "Application"      = @("Application")
    "PowerShell"       = @("Microsoft-Windows-PowerShell/Operational")
    "Windows Defender" = @("Microsoft-Windows-Windows Defender/Operational")
    "WLAN / Netzwerk"  = @("Microsoft-Windows-WLAN-AutoConfig/Operational")
    "Task Scheduler"   = @("Microsoft-Windows-TaskScheduler/Operational")
    "Remote Desktop"   = @(
        "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational",
        "Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational"
    )
    "Druckerdienst"    = @("Microsoft-Windows-PrintService/Operational")
    "BitLocker"        = @("Microsoft-Windows-BitLocker/BitLocker Management")
    "SMB Client"       = @("Microsoft-Windows-SmbClient/Connectivity")
    "Eigene IDs"       = @("__CUSTOM__")
}

# Index-Liste: CLB-Position -> Katalog-Eintrag (vermeidet fehleranfaellige String-Suche)
$script:visibleEvents = [System.Collections.Generic.List[PSObject]]::new()

# Funktion: Liste befuellen
# Kurzlabel für Log-Namen (damit die CLB-Zeile nicht überläuft)
function Get-LogShortLabel($logName) {
    switch -Wildcard ($logName) {
        "System"                                                             { return "System" }
        "Security"                                                           { return "Security" }
        "Application"                                                        { return "Application" }
        "Microsoft-Windows-PowerShell/Operational"                           { return "PowerShell" }
        "Microsoft-Windows-Windows Defender/Operational"                     { return "Defender" }
        "Microsoft-Windows-WLAN-AutoConfig/Operational"                      { return "WLAN" }
        "Microsoft-Windows-TaskScheduler/Operational"                        { return "TaskSched" }
        "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational" { return "RDP-LSM" }
        "Microsoft-Windows-TerminalServices-RemoteConnectionManager/*"       { return "RDP-RCM" }
        "Microsoft-Windows-PrintService/Operational"                         { return "Print" }
        "Microsoft-Windows-BitLocker/BitLocker Management"                   { return "BitLocker" }
        "Microsoft-Windows-SmbClient/Connectivity"                           { return "SMB" }
        default                                                              { return ($logName -split '/')[-1] }
    }
}

function Update-EventList {
    $clbEvents.Items.Clear()
    $script:visibleEvents.Clear()
    $filterKat = $cbKat.SelectedItem
    $filterLog = $cbLog.SelectedItem

    # Bei Scan-Ergebnis-Ansichten nach Log und ID sortieren - sonst Katalog-Reihenfolge
    $isScanView = $filterLog -in @(">> Gefunden (alle Scan-Ergebnisse)", "◈ Erkannt auf Computer", "= Dokumentiert (Manifest)")
    $iterList = if ($isScanView) {
        $eventCatalog | Sort-Object Log, ID
    } else {
        $eventCatalog
    }

    foreach ($ev in $iterList) {
        $showKat = ($filterKat -eq "Alle Kategorien") -or ($ev.Category -eq $filterKat)

        # Sonderfälle für die Bereichs-Filter
        $showLog = switch ($filterLog) {
            "Alle Bereiche"                       { $true }
            ">> Gefunden (alle Scan-Ergebnisse)"  { $ev.Category -eq "Erkannt" -or $ev.Category -eq "Dokumentiert" }
            "★ Empfohlen (kuratiert)"             { $ev.Category -ne "Erkannt" -and $ev.Category -ne "Eigene" -and $ev.Category -ne "Dokumentiert" }
            "◈ Erkannt auf Computer"              { $ev.Category -eq "Erkannt" }
            "= Dokumentiert (Manifest)"          { $ev.Category -eq "Dokumentiert" }
            "Eigene IDs"                          { $ev.Category -eq "Eigene" }
            default {
                if ($script:logGroupMap.ContainsKey($filterLog)) {
                    $script:logGroupMap[$filterLog] -contains $ev.Log
                } else { $false }
            }
        }

        $searchTerm = $script:searchFilter.Trim().ToLower()
        $matchSearch = $searchTerm -eq "" -or
                       ($ev.ID.ToString()         -like "*$searchTerm*") -or
                       ($ev.Desc.ToLower()         -like "*$searchTerm*") -or
                       ($ev.Category.ToLower()     -like "*$searchTerm*") -or
                       ($ev.Log.ToLower()          -like "*$searchTerm*")

        if ($showKat -and $showLog -and $matchSearch) {
            $logShort = Get-LogShortLabel $ev.Log
            $display  = ("{0}  ID {1,-6} [{2,-10}] [{3,-12}]  {4}" -f $ev.Icon, $ev.ID, $logShort, $ev.Category, $ev.Desc)
            $clbEvents.Items.Add($display) | Out-Null
            $script:visibleEvents.Add($ev)
        }
    }
    $lblCatCount.Text = "$($script:visibleEvents.Count) Einträge"
}

# Tooltips: Event-Auswahl-Sektion
$toolTip.SetToolTip($cbKat,
    "Filtert die Event-Liste nach Schweregrad/Kategorie.`n`n" +
    "Kritisch / Fehler / Warnung / Information = Windows-Schweregrade`n" +
    "Erkannt = auf diesem Computer gefundene Events`n" +
    "Eigene  = manuell hinzugefügte Event-IDs")

$toolTip.SetToolTip($cbLog,
    "Filtert die Event-Liste nach Log-Bereich.`n`n" +
    ">> Gefunden = alle erkannten und dokumentierten Scan-Ergebnisse`n" +
    "★ Empfohlen = kuratierte, praxiserprobte Event-IDs`n" +
    "◈ Erkannt   = aktiv auf dem Computer vorhandene Events`n" +
    "= Dokumentiert = aus Provider-Manifesten (Manifest-Scan)")

$toolTip.SetToolTip($btnAll,
    "Alle sichtbaren Events in der Liste ankreuzen.")

$toolTip.SetToolTip($btnNone,
    "Alle Häkchen in der Liste entfernen.")

$toolTip.SetToolTip($clbEvents,
    "Event-Auswahlliste: Gewünschte Event-IDs ankreuzen.`n" +
    "Ein Klick auf einen Eintrag zeigt unten die vollständige Beschreibung.`n`n" +
    "Format: [Icon]  ID [Nummer]  [Log]  [Kategorie]  Beschreibung")

# Filter-Events verdrahten
$cbKat.Add_SelectedIndexChanged({ Update-EventList })
$cbLog.Add_SelectedIndexChanged({ Update-EventList })
$txtSearch.Add_TextChanged({
    $script:searchFilter = $txtSearch.Text
    Update-EventList
})
$btnSearchClear.Add_Click({
    $txtSearch.Clear()
    $script:searchFilter = ""
    Update-EventList
})
$btnAll.Add_Click({ for ($i=0; $i -lt $clbEvents.Items.Count; $i++) { $clbEvents.SetItemChecked($i, $true) } })
$btnNone.Add_Click({ for ($i=0; $i -lt $clbEvents.Items.Count; $i++) { $clbEvents.SetItemChecked($i, $false) } })

# Re-Scan Handler
$btnRescan.Add_Click({
    $target = $txtComputer.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($target)) { $target = $env:COMPUTERNAME }
    $withManifest = $chkManifest.Checked

    # Vorhandene Erkannt- und Dokumentiert-Einträge entfernen
    $toRemove = @($eventCatalog | Where-Object { $_.Category -eq "Erkannt" -or $_.Category -eq "Dokumentiert" })
    foreach ($r in $toRemove) { [void]$eventCatalog.Remove($r) }

    $progUI2 = New-ProgressForm -ComputerName $target
    $progUI2.Form.Show()
    $progUI2.Form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $cred  = Get-FormCredential

        # Phase 1: Schnell-Scan (immer)
        $disc  = Invoke-ComputerEventScan -Computer $target -ProgressUI $progUI2 `
                     -MaxPerLog 15 -MaxLogs 120 -Credential $cred
        $added = Merge-DiscoveredEvents -Discovered $disc

        # Phase 2: Manifest-Scan (nur wenn Checkbox aktiv)
        $addedDocs2 = 0
        if ($withManifest -and -not $script:scanCancelled) {
            $relLogs = @($disc | Select-Object -ExpandProperty Log -Unique)
            if ($relLogs.Count -gt 0) {
                $docs2 = Invoke-DocumentedIDsScan -Computer $target -ProgressUI $progUI2 -LogNames $relLogs -Credential $cred
                $addedDocs2 = Merge-DiscoveredEvents -Discovered $docs2
            }
        }

        $progUI2.Form.Close(); $progUI2.Form.Dispose()

        if ($script:scanCancelled) {
            $lblStatus.ForeColor = $clrMuted
            $lblStatus.Text = "Scan abgebrochen."
        } else {
            $suffix = if ($withManifest) { " + $addedDocs2 dokumentierte IDs" } else { "" }
            $lblStatus.ForeColor = $clrSuccess
            $lblStatus.Text = "✔ Scan: $added erkannte IDs$suffix aus $($script:discoveredLogs.Count) Logs."
            $cbLog.SelectedItem = ">> Gefunden (alle Scan-Ergebnisse)"
        }
    } catch {
        try { $progUI2.Form.Close(); $progUI2.Form.Dispose() } catch {}
        $lblStatus.ForeColor = $clrWarn
        $lblStatus.Text = "⚠  Scan fehlgeschlagen: $($_.Exception.Message)"
    }

    Update-EventList
})

# Wenn der Startup-Scan etwas gefunden hat: Filter auf "Gefunden" voreinstellen
$hasFindings = ($eventCatalog | Where-Object { $_.Category -eq "Erkannt" -or $_.Category -eq "Dokumentiert" } | Select-Object -First 1)
if ($hasFindings) {
    $cbLog.SelectedItem = ">> Gefunden (alle Scan-Ergebnisse)"   # SelectedIndexChanged löst Update-EventList aus
} else {
    Update-EventList
}

# ── Sektion: Eigene Event-IDs ─────────────────────────────────
$pnlCustom = New-SectionPanel 15 625 885 85 "+  Eigene Event-ID hinzufügen"
$formMain.Controls.Add($pnlCustom)

# Event-ID
$pnlCustom.Controls.Add((New-Label "Event-ID:" 10 35 70 20))
$txtCustomID = New-Object System.Windows.Forms.TextBox
$txtCustomID.Location = New-Object System.Drawing.Point(82, 33)
$txtCustomID.Size     = New-Object System.Drawing.Size(80, 22)
$txtCustomID.Font     = $fontNormal
$pnlCustom.Controls.Add($txtCustomID)

# Protokoll (freie Auswahl: Dropdown mit Vorschlägen + editierbar)
$pnlCustom.Controls.Add((New-Label "Log/Protokoll:" 180 35 90 20))
$cbCustomLog = New-Object System.Windows.Forms.ComboBox
$cbCustomLog.Location      = New-Object System.Drawing.Point(275, 33)
$cbCustomLog.Size          = New-Object System.Drawing.Size(340, 22)
$cbCustomLog.DropDownStyle = "DropDown"   # editierbar!
$cbCustomLog.Font          = $fontNormal
# Liste aller bekannten/nützlichen Logs
@(
    "System", "Security", "Application", "Setup",
    "Microsoft-Windows-PowerShell/Operational",
    "Windows PowerShell",
    "Microsoft-Windows-Windows Defender/Operational",
    "Microsoft-Windows-WLAN-AutoConfig/Operational",
    "Microsoft-Windows-TaskScheduler/Operational",
    "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational",
    "Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational",
    "Microsoft-Windows-PrintService/Operational",
    "Microsoft-Windows-BitLocker/BitLocker Management",
    "Microsoft-Windows-SmbClient/Connectivity",
    "Microsoft-Windows-GroupPolicy/Operational",
    "Microsoft-Windows-AppLocker/EXE and DLL",
    "Microsoft-Windows-Kernel-PnP/Configuration",
    "Microsoft-Windows-Sysmon/Operational"
) | ForEach-Object { $cbCustomLog.Items.Add($_) | Out-Null }
$cbCustomLog.Text = "System"
$pnlCustom.Controls.Add($cbCustomLog)

# Beschreibung (optional)
$pnlCustom.Controls.Add((New-Label "Beschreibung:" 10 60 90 20))
$txtCustomDesc = New-Object System.Windows.Forms.TextBox
$txtCustomDesc.Location = New-Object System.Drawing.Point(100, 58)
$txtCustomDesc.Size     = New-Object System.Drawing.Size(515, 22)
$txtCustomDesc.Font     = $fontNormal
$txtCustomDesc.Text     = ""
$pnlCustom.Controls.Add($txtCustomDesc)

# Hinzufügen-Button
$btnCustomAdd = New-StyledButton "+ Hinzufügen" 630 31 120 26 $true
$pnlCustom.Controls.Add($btnCustomAdd)

# Info-Label
$lblCustomInfo = New-Label "Das Log lässt sich auch frei eingeben (siehe Ereignisanzeige → Eigenschaften)" 630 61 245 20 $false $true
$lblCustomInfo.Font = $fontSmall
$pnlCustom.Controls.Add($lblCustomInfo)

# Klick-Handler
$btnCustomAdd.Add_Click({
    $idRaw  = $txtCustomID.Text.Trim()
    $logVal = $cbCustomLog.Text.Trim()
    $desc   = $txtCustomDesc.Text.Trim()

    if ($idRaw -eq "" -or -not ($idRaw -match '^\d+$')) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte eine gültige Event-ID (nur Ziffern) eingeben.",
            "Ungültige Eingabe", "OK", "Warning") | Out-Null
        $txtCustomID.Focus(); return
    }
    if ($logVal -eq "") {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte ein Protokoll/Log angeben.",
            "Fehlende Eingabe", "OK", "Warning") | Out-Null
        $cbCustomLog.Focus(); return
    }

    # Doppelt?
    $id = [int]$idRaw
    $dup = $eventCatalog | Where-Object { $_.ID -eq $id -and $_.Log -eq $logVal } | Select-Object -First 1
    if ($dup) {
        [System.Windows.Forms.MessageBox]::Show(
            "ID $id in Log '$logVal' ist bereits im Katalog.",
            "Bereits vorhanden", "OK", "Information") | Out-Null
        return
    }

    if ($desc -eq "") { $desc = "(Eigene ID – keine Beschreibung hinterlegt)" }

    $newEv = [PSCustomObject]@{
        ID       = $id
        Log      = $logVal
        Category = "Eigene"
        Icon     = "★"
        Desc     = $desc
    }
    $eventCatalog.Add($newEv)

    # Liste aktualisieren und auf neuen Eintrag schalten
    Update-EventList

    # Neuen Eintrag automatisch ankreuzen, wenn sichtbar
    for ($i = 0; $i -lt $script:visibleEvents.Count; $i++) {
        if ($script:visibleEvents[$i] -eq $newEv) {
            $clbEvents.SetItemChecked($i, $true)
            $clbEvents.TopIndex = [Math]::Max(0, $i - 3)
            $clbEvents.SelectedIndex = $i
            break
        }
    }

    # Felder zurücksetzen
    $txtCustomID.Text   = ""
    $txtCustomDesc.Text = ""
    $txtCustomID.Focus()
})

# Beschreibungs-Label
$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Location  = New-Object System.Drawing.Point(15, 718)
$lblDesc.Size      = New-Object System.Drawing.Size(885, 22)
$lblDesc.Font      = $fontSmall
$lblDesc.ForeColor = $clrMuted
$lblDesc.Text      = "ℹ  Hover über Event-ID für Details"
$formMain.Controls.Add($lblDesc)

$clbEvents.Add_SelectedIndexChanged({
    $idx = $clbEvents.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:visibleEvents.Count) {
        $ev = $script:visibleEvents[$idx]
        $lblDesc.Text = "ℹ  ID $($ev.ID)  ·  $($ev.Log)  ·  $($ev.Category)  →  $($ev.Desc)"
        $lblDesc.ForeColor = $clrAccent
    }
})

# ── Aktions-Buttons ───────────────────────────────────────────
$btnAbfragen = New-StyledButton "Abfragen" 700 748 200 42 $true
$btnAbfragen.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$formMain.Controls.Add($btnAbfragen)

$btnBeenden = New-StyledButton "Beenden" 15 748 100 42 $false
$formMain.Controls.Add($btnBeenden)
$btnBeenden.Add_Click({ $formMain.Close() })

$btnXPath = New-StyledButton "XPath-Abfrage" 560 748 130 42 $false
$formMain.Controls.Add($btnXPath)

# Tooltips: Eigene Event-ID Sektion
$toolTip.SetToolTip($txtCustomID,
    "Numerische Event-ID, die hinzugefügt werden soll.`n" +
    "Beispiel: 4625 (fehlgeschlagene Anmeldung)")

$toolTip.SetToolTip($cbCustomLog,
    "Windows-Ereignisprotokoll, zu dem die ID gehört.`n" +
    "Auswahl wählen oder eigenen Log-Namen eingeben.`n`n" +
    "Den Log-Namen finden Sie in der Ereignisanzeige`n" +
    "unter Protokolleigenschaften.")

$toolTip.SetToolTip($txtCustomDesc,
    "Optionale Beschreibung für diese Event-ID.`n" +
    "Wird in der Event-Liste angezeigt.")

$toolTip.SetToolTip($btnCustomAdd,
    "Eigene Event-ID zur Auswahlliste hinzufügen.`n" +
    "Der Eintrag erscheint in der Kategorie 'Eigene'.")

# Tooltips: Aktions-Buttons
$toolTip.SetToolTip($btnAbfragen,
    "Windows-Ereignisprotokoll mit allen angekreuzten Event-IDs abfragen.`n" +
    "Das Ergebnis wird in einem neuen Fenster angezeigt.`n`n" +
    "Tipp: Zeitraum und Max. Einträge oben anpassen,`n" +
    "um die Abfrage zu beschleunigen.")

$toolTip.SetToolTip($btnXPath,
    "XPath-Direktabfrage öffnen: Freies XPath-Filterfeld`n" +
    "für erweiterte Filterung des Windows-Ereignisprotokolls.`n`n" +
    "Beispiel: *[System[EventID=4625]]")

$toolTip.SetToolTip($btnBeenden,
    "Programm beenden.")

$btnXPath.Add_Click({
    # ── XPath-Dialog ──────────────────────────────────────────────
    $fxp = New-Object System.Windows.Forms.Form
    $fxp.Text            = "XPath-Direktabfrage"
    $fxp.Size            = New-Object System.Drawing.Size(710, 370)
    $fxp.StartPosition   = "CenterScreen"
    $fxp.FormBorderStyle = "FixedDialog"
    $fxp.MaximizeBox     = $false
    $fxp.BackColor       = $clrBg
    $fxp.Font            = $fontNormal

    $pnlXH = New-Object System.Windows.Forms.Panel
    $pnlXH.Dock = "Top"; $pnlXH.Height = 48; $pnlXH.BackColor = $clrAccent
    $lblXH = New-Object System.Windows.Forms.Label
    $lblXH.Text = "  XPath-Direktabfrage"; $lblXH.Dock = "Fill"
    $lblXH.Font = $fontTitle; $lblXH.ForeColor = [System.Drawing.Color]::White; $lblXH.TextAlign = "MiddleLeft"
    $pnlXH.Controls.Add($lblXH); $fxp.Controls.Add($pnlXH)

    # Log-Name
    $fxp.Controls.Add((New-Label "Log-Name:" 15 66 72 20 $true))
    $cbXLog = New-Object System.Windows.Forms.ComboBox
    $cbXLog.Location = New-Object System.Drawing.Point(92, 64)
    $cbXLog.Size = New-Object System.Drawing.Size(345, 22); $cbXLog.Font = $fontNormal
    $cbXLog.DropDownStyle = "DropDown"
    @("System","Security","Application",
      "Microsoft-Windows-PowerShell/Operational",
      "Microsoft-Windows-Windows Defender/Operational",
      "Microsoft-Windows-TaskScheduler/Operational",
      "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational",
      "Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational"
    ) | ForEach-Object { $cbXLog.Items.Add($_) | Out-Null }
    $cbXLog.Text = "System"; $fxp.Controls.Add($cbXLog)

    # Computer (aus Hauptformular vorbelegen)
    $fxp.Controls.Add((New-Label "Computer:" 455 66 68 20 $true))
    $txtXComp = New-Object System.Windows.Forms.TextBox
    $txtXComp.Location = New-Object System.Drawing.Point(528, 64); $txtXComp.Size = New-Object System.Drawing.Size(160, 22)
    $txtXComp.Text = $txtComputer.Text.Trim(); $fxp.Controls.Add($txtXComp)

    # Max Events
    $fxp.Controls.Add((New-Label "Max. Einträge:" 15 98 95 20))
    $cbXMax = New-Object System.Windows.Forms.ComboBox
    $cbXMax.Location = New-Object System.Drawing.Point(115, 96); $cbXMax.Size = New-Object System.Drawing.Size(80, 22)
    $cbXMax.DropDownStyle = "DropDownList"
    @(25,50,100,250,500) | ForEach-Object { $cbXMax.Items.Add($_) | Out-Null }
    $cbXMax.SelectedIndex = 2; $fxp.Controls.Add($cbXMax)

    # XPath-Filter
    $fxp.Controls.Add((New-Label "XPath-Filter:" 15 130 85 20 $true))
    $txtXPath2 = New-Object System.Windows.Forms.TextBox
    $txtXPath2.Location  = New-Object System.Drawing.Point(15, 152)
    $txtXPath2.Size      = New-Object System.Drawing.Size(668, 100)
    $txtXPath2.Multiline = $true; $txtXPath2.Font = $fontMono; $txtXPath2.ScrollBars = "Vertical"
    $txtXPath2.Text = "*[System[(Level=1 or Level=2) and TimeCreated[timediff(@SystemTime) <= 86400000]]]"
    $fxp.Controls.Add($txtXPath2)

    $lblXEx = New-Label "z.B.: *[System[EventID=4625 and TimeCreated[timediff(@SystemTime) <= 3600000]]]" 15 258 668 18 $false $true
    $lblXEx.Font = $fontSmall; $fxp.Controls.Add($lblXEx)

    $btnXRun    = New-StyledButton "Abfragen" 455 285 115 32 $true
    $btnXClose  = New-StyledButton "Abbrechen"   578 285 100 32 $false
    $lblXStat   = New-Label "" 15 292 432 22 $false $true
    $fxp.Controls.Add($btnXRun); $fxp.Controls.Add($btnXClose); $fxp.Controls.Add($lblXStat)
    $btnXClose.Add_Click({ $fxp.Close() })

    $btnXRun.Add_Click({
        $xLog   = $cbXLog.Text.Trim()
        $xXPath = $txtXPath2.Text.Trim()
        $xComp  = $txtXComp.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($xComp)) { $xComp = $env:COMPUTERNAME }
        $xMax   = [int]$cbXMax.SelectedItem

        if ([string]::IsNullOrWhiteSpace($xLog) -or [string]::IsNullOrWhiteSpace($xXPath)) {
            $lblXStat.ForeColor = $clrWarn; $lblXStat.Text = "⚠  Log-Name und XPath-Filter sind erforderlich."; return
        }
        $lblXStat.ForeColor = $clrMuted; $lblXStat.Text = "Abfrage läuft..."; $fxp.Refresh()

        try {
            $xCred = Get-FormCredential
            $xqp = @{ LogName = $xLog; FilterXPath = $xXPath; MaxEvents = $xMax; ComputerName = $xComp; ErrorAction = 'Stop' }
            if ($xCred) { $xqp.Credential = $xCred }
            $xRaw = Get-WinEvent @xqp

            $xSorted = @($xRaw | ForEach-Object {
                $r = $_
                $msg   = if ($r.Message) { $r.Message } else { "(keine Nachricht)" }
                $short = ($msg -replace "`r`n|`n", " ")
                if ($short.Length -gt 300) { $short = $short.Substring(0,300) + "..." }
                $typ = switch ($r.LevelDisplayName) {
                    "Fehler" { "Error" } "Error" { "Error" } "Kritisch" { "Critical" } "Critical" { "Critical" }
                    "Warnung" { "Warning" } "Warning" { "Warning" }
                    "Information" { "Information" } "Informationen" { "Information" }
                    default { "$($r.LevelDisplayName)" }
                }
                [PSCustomObject]@{
                    Computer  = $xComp
                    Zeit      = $r.TimeCreated; Typ = $typ
                    Protokoll = Get-LogShortLabel $xLog; LogVoll = $xLog
                    EventID   = $r.Id; Quelle = $r.ProviderName
                    Kategorie = "XPath"; Beschr = $xXPath.Substring(0,[Math]::Min(60,$xXPath.Length))
                    Nachricht = $short
                }
            } | Sort-Object Zeit -Descending)

            if ($xSorted.Count -eq 0) {
                $lblXStat.ForeColor = $clrWarn; $lblXStat.Text = "⚠  Keine Ereignisse gefunden."; return
            }
            $lblXStat.ForeColor = $clrSuccess
            $lblXStat.Text = "✔  $($xSorted.Count) Ereignisse – Ausgabe wird geöffnet..."

            # ── XPath Ausgabe-Formular ─────────────────────────────────
            $fxOut = New-Object System.Windows.Forms.Form
            $fxOut.Text          = "XPath-Ergebnisse  –  $($xSorted.Count) Einträge  |  $xLog  |  $xComp"
            $fxOut.Size          = New-Object System.Drawing.Size(1150, 720)
            $fxOut.StartPosition = "CenterScreen"; $fxOut.BackColor = $clrBg; $fxOut.Font = $fontNormal

            $pnlXT = New-Object System.Windows.Forms.Panel
            $pnlXT.Dock = "Top"; $pnlXT.Height = 56; $pnlXT.BackColor = $clrAccent
            $lblXT = New-Object System.Windows.Forms.Label
            $lblXT.Text = "  XPath-Ergebnisse  ·  $($xSorted.Count) Einträge"
            $lblXT.Dock = "Fill"; $lblXT.Font = $fontTitle; $lblXT.ForeColor = [System.Drawing.Color]::White
            $lblXT.TextAlign = "MiddleLeft"; $pnlXT.Controls.Add($lblXT); $fxOut.Controls.Add($pnlXT)

            $pnlXF2 = New-Object System.Windows.Forms.Panel
            $pnlXF2.Location = New-Object System.Drawing.Point(10,64); $pnlXF2.Size = New-Object System.Drawing.Size(1115,44)
            $pnlXF2.BackColor = $clrPanel
            $pnlXF2.Add_Paint({ param($s,$e); $e.Graphics.DrawRectangle((New-Object System.Drawing.Pen($clrBorder,1)),0,0,$s.Width-1,$s.Height-1) })
            $lblXF2 = New-Label "Filter:" 8 12 60 20 $true; $pnlXF2.Controls.Add($lblXF2)
            $txtXF2 = New-Object System.Windows.Forms.TextBox
            $txtXF2.Location = New-Object System.Drawing.Point(70,10); $txtXF2.Size = New-Object System.Drawing.Size(280,22)
            $txtXF2.Font = $fontNormal; $pnlXF2.Controls.Add($txtXF2)
            $lblXC2 = New-Label "" 700 12 280 20 $false $true; $pnlXF2.Controls.Add($lblXC2)
            $btnXE2 = New-StyledButton "CSV-Export" 990 8 110 28 $false; $pnlXF2.Controls.Add($btnXE2)
            $fxOut.Controls.Add($pnlXF2)

            # DGV
            $xdgv = New-Object System.Windows.Forms.DataGridView
            $xdgv.Location = New-Object System.Drawing.Point(10,114); $xdgv.Size = New-Object System.Drawing.Size(1115,500)
            $xdgv.BackgroundColor = $clrPanel; $xdgv.GridColor = $clrBorder; $xdgv.BorderStyle = "None"
            $xdgv.Font = $fontSmall; $xdgv.DefaultCellStyle.Font = $fontSmall
            $xdgv.DefaultCellStyle.ForeColor = $clrText; $xdgv.DefaultCellStyle.BackColor = $clrPanel
            $xdgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(220,225,255)
            $xdgv.DefaultCellStyle.SelectionForeColor = $clrText
            $xdgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235,235,248)
            $xdgv.ColumnHeadersDefaultCellStyle.ForeColor = $clrAccent
            $xdgv.ColumnHeadersDefaultCellStyle.Font = $fontBold
            $xdgv.ColumnHeadersHeight = 28; $xdgv.RowHeadersVisible = $false
            $xdgv.AllowUserToAddRows = $false; $xdgv.AllowUserToDeleteRows = $false
            $xdgv.ReadOnly = $true; $xdgv.SelectionMode = "FullRowSelect"
            $xdgv.RowTemplate.Height = 22; $xdgv.EnableHeadersVisualStyles = $false
            $xdgv.ScrollBars = "Both"; $fxOut.Controls.Add($xdgv)

            foreach ($xColDef in @(
                @{N="Computer";W=100},@{N="Zeit";W=135},@{N="Typ";W=80},
                @{N="Protokoll";W=85},@{N="EventID";W=65},@{N="Quelle";W=160},@{N="Nachricht";Fill=$true}
            )) {
                $xc = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
                $xc.Name = $xColDef.N; $xc.HeaderText = $xColDef.N; $xc.SortMode = "Automatic"
                if ($xColDef.Fill) { $xc.AutoSizeMode = "Fill" } else { $xc.Width = $xColDef.W }
                $xdgv.Columns.Add($xc) | Out-Null
            }
            $xdgv.Add_CellFormatting({ param($s,$e)
                if ($e.RowIndex -lt 0) { return }
                $t = $s.Rows[$e.RowIndex].Cells["Typ"].Value
                $e.CellStyle.BackColor = switch ($t) {
                    "Critical" { [System.Drawing.Color]::FromArgb(255,225,225) }
                    "Error"    { [System.Drawing.Color]::FromArgb(255,240,240) }
                    "Warning"  { [System.Drawing.Color]::FromArgb(255,252,235) }
                    default    { $clrPanel }
                }
            })

            $script:xSortedRef = $xSorted
            foreach ($r in $xSorted) {
                $xdgv.Rows.Add($r.Computer, $r.Zeit.ToString("dd.MM.yyyy HH:mm:ss"), $r.Typ,
                    $r.Protokoll, $r.EventID, $r.Quelle, $r.Nachricht) | Out-Null
            }
            $lblXC2.Text = "Einträge: $($xSorted.Count)"

            $txtXF2.Add_TextChanged({
                $f2 = $txtXF2.Text.Trim()
                $xdgv.Rows.Clear()
                $xFil = if ($f2 -eq "") { $script:xSortedRef } else {
                    $script:xSortedRef | Where-Object { $_.EventID -like "*$f2*" -or $_.Quelle -like "*$f2*" -or $_.Nachricht -like "*$f2*" }
                }
                foreach ($r in $xFil) {
                    $xdgv.Rows.Add($r.Computer, $r.Zeit.ToString("dd.MM.yyyy HH:mm:ss"), $r.Typ,
                        $r.Protokoll, $r.EventID, $r.Quelle, $r.Nachricht) | Out-Null
                }
                $lblXC2.Text = "Einträge: $($xFil.Count) / $($script:xSortedRef.Count)"
            })

            $btnXE2.Add_Click({
                $dlgX = New-Object System.Windows.Forms.SaveFileDialog
                $dlgX.Filter   = "CSV-Datei (*.csv)|*.csv"
                $dlgX.FileName = "XPath_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
                if ($dlgX.ShowDialog() -eq "OK") {
                    $script:xSortedRef | Export-Csv -Path $dlgX.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
                    [System.Windows.Forms.MessageBox]::Show("Export: $($dlgX.FileName)", "OK","OK","Information") | Out-Null
                }
            })

            $fxOut.ShowDialog() | Out-Null

        } catch [System.Exception] {
            if ($_.Exception.Message -match "No events were found|Es wurden keine Ereignisse gefunden") {
                $lblXStat.ForeColor = $clrWarn; $lblXStat.Text = "⚠  Keine Ereignisse gefunden."
            } else {
                $lblXStat.ForeColor = $clrWarn; $lblXStat.Text = "⚠  $($_.Exception.Message)"
            }
        }
    })

    $fxp.ShowDialog() | Out-Null
})

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(130, 790)
$lblStatus.Size      = New-Object System.Drawing.Size(555, 32)
$lblStatus.Font      = $fontSmall
$lblStatus.ForeColor = $clrMuted
if ($script:startupScanResult) {
    $lblStatus.Text = "✔ $($script:startupScanResult)"
    $lblStatus.ForeColor = $clrSuccess
} else {
    $lblStatus.Text = "Bitte Events auswählen und 'Abfragen' klicken..."
}
$formMain.Controls.Add($lblStatus)

# ════════════════════════════════════════════════════════════
#  ABFRAGE-LOGIK
# ════════════════════════════════════════════════════════════
$btnAbfragen.Add_Click({
    # Ausgewaehlte Events per Index direkt aus der Mapping-Liste holen
    $checkedIDs = @()
    for ($i = 0; $i -lt $clbEvents.Items.Count; $i++) {
        if ($clbEvents.GetItemChecked($i) -and $i -lt $script:visibleEvents.Count) {
            $checkedIDs += $script:visibleEvents[$i]
        }
    }

    if ($checkedIDs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Bitte mindestens eine Event-ID auswählen.",
            "Keine Auswahl",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    # Zeitraum berechnen
    $zeitText = $cbZeit.SelectedItem
    $startTime = switch ($zeitText) {
        "Letzte 1 Stunde"   { (Get-Date).AddHours(-1) }
        "Letzte 6 Stunden"  { (Get-Date).AddHours(-6) }
        "Letzte 24 Stunden" { (Get-Date).AddHours(-24) }
        "Letzte 7 Tage"     { (Get-Date).AddDays(-7) }
        "Letzte 30 Tage"    { (Get-Date).AddDays(-30) }
        default             { $null }
    }

    $maxCount      = [int]$cbMax.SelectedItem
    $computers     = @($txtComputer.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    if ($computers.Count -eq 0) { $computers = @($env:COMPUTERNAME) }
    $cred          = Get-FormCredential
    $computer      = $computers[0]   # Compat für LiveTimer / FormOut.Text
    $computerLabel = if ($computers.Count -eq 1) { $computers[0] } else { "$($computers.Count) Computer" }

    $lblStatus.ForeColor = $clrMuted
    $lblStatus.Text      = "Abfrage läuft..."
    $formMain.Refresh()

    # Events abfragen – pro Computer, gruppiert nach Log
    $allResults = [System.Collections.Generic.List[PSObject]]::new()
    $errors     = @()
    $diag       = [System.Collections.Generic.List[string]]::new()

    foreach ($computer in $computers) {
        # Pro Log-Name alle selektierten IDs sammeln, damit wir nur 1x pro Log abfragen
        $byLog = $checkedIDs | Group-Object -Property Log

        foreach ($grp in $byLog) {
            $logName = $grp.Name
            $allIds  = @($grp.Group | Select-Object -ExpandProperty ID -Unique)

            # Lookup: ID -> Katalog-Eintrag (für Kategorie/Desc im Ergebnis)
            $idLookup = @{}
            foreach ($ev in $grp.Group) { $idLookup[$ev.ID] = $ev }

            # WICHTIG: Get-WinEvent -FilterHashtable unterstützt MAXIMAL 22 IDs pro Aufruf!
            $batchSize = 22
            $hitCountThisLog = 0
            $batchErrors     = @()

            for ($offset = 0; $offset -lt $allIds.Count; $offset += $batchSize) {
                $end       = [Math]::Min($offset + $batchSize - 1, $allIds.Count - 1)
                $idsBatch  = $allIds[$offset..$end]

                $filter = @{ LogName = $logName; ID = $idsBatch }
                if ($startTime) { $filter.StartTime = $startTime }

                try {
                    $queryParams = @{ ComputerName = $computer; FilterHashtable = $filter; MaxEvents = $maxCount; ErrorAction = 'Stop' }
                    if ($cred) { $queryParams.Credential = $cred }
                    $winEvents = Get-WinEvent @queryParams

                    foreach ($r in $winEvents) {
                        $hitCountThisLog++
                        $meta = $idLookup[[int]$r.Id]
                        $msg  = if ($r.Message) { $r.Message } else { "(keine Nachricht)" }
                        $short = ($msg -replace "`r`n|`n", " ")
                        if ($short.Length -gt 300) { $short = $short.Substring(0, 300) + "..." }

                        $typ = switch ($r.LevelDisplayName) {
                            "Fehler"          { "Error" }
                            "Error"           { "Error" }
                            "Kritisch"        { "Critical" }
                            "Critical"        { "Critical" }
                            "Warnung"         { "Warning" }
                            "Warning"         { "Warning" }
                            "Information"     { "Information" }
                            "Informationen"   { "Information" }
                            "Ausführlich"     { "Verbose" }
                            "Verbose"         { "Verbose" }
                            default           { "$($r.LevelDisplayName)" }
                        }

                        $allResults.Add([PSCustomObject]@{
                            Computer  = $computer
                            Zeit      = $r.TimeCreated
                            Typ       = $typ
                            Protokoll = Get-LogShortLabel $logName
                            LogVoll   = $logName
                            EventID   = $r.Id
                            Quelle    = $r.ProviderName
                            Kategorie = if ($meta) { $meta.Category } else { "" }
                            Beschr    = if ($meta) { $meta.Desc }     else { "" }
                            Nachricht = $short
                        })
                    }
                } catch [System.Exception] {
                    if ($_.Exception.Message -match "No events were found|Es wurden keine Ereignisse gefunden") {
                        # nichts tun
                    } else {
                        $batchErrors += "$($_.Exception.Message)"
                    }
                }
            }

            # Diagnose-Eintrag pro Log
            $logShort = Get-LogShortLabel $logName
            $diag.Add("[$computer] $logShort  ·  $($allIds.Count) ID(s)  ·  $hitCountThisLog Treffer")

            if ($batchErrors.Count -gt 0) {
                $errors += "[$computer] '$logName': " + ($batchErrors | Select-Object -Unique -First 1)
            }
        }
    }

    # Ergebnisse sortieren
    $sorted = $allResults | Sort-Object Computer, Zeit -Descending

    if ($errors.Count -gt 0 -and $sorted.Count -eq 0) {
        $lblStatus.ForeColor = $clrWarn
        $lblStatus.Text = "⚠  Fehler bei allen Abfragen. Ggf. als Admin ausführen."
        [System.Windows.Forms.MessageBox]::Show(
            ($errors -join "`n"),
            "Abfragefehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    # Wenn keine Treffer: aussagekräftigen Hinweis mit Diagnose anzeigen
    if ($sorted.Count -eq 0) {
        $lblStatus.ForeColor = $clrWarn
        $lblStatus.Text = "⚠  Keine Ereignisse im gewählten Zeitraum gefunden."
        $diagText = @"
Keine Ereignisse im Zeitraum '$zeitText' gefunden.

Mögliche Gründe:
 - Die ausgewählten IDs sind dokumentiert, aber nicht aufgetreten
 - Der Zeitraum ist zu eng (versuche '7 Tage' oder 'Alles')
 - Adminrechte fehlen für einzelne Logs

Abfrage-Statistik:
$($diag -join "`n")
"@
        if ($errors.Count -gt 0) {
            $diagText += "`n`nFehler:`n" + ($errors -join "`n")
        }
        [System.Windows.Forms.MessageBox]::Show(
            $diagText,
            "Keine Ereignisse",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $lblStatus.ForeColor = $clrSuccess
    $lblStatus.Text = "✔  $($sorted.Count) Ereignis(se) gefunden – Ausgabe wird geöffnet..."

    # ════════════════════════════════════════════════════════
    #  FORMULAR 2 – Ausgabe
    # ════════════════════════════════════════════════════════
    $formOut = New-Object System.Windows.Forms.Form
    $formOut.Text           = "Event-Ergebnisse  –  $($sorted.Count) Einträge  |  $zeitText  |  $computerLabel"
    $formOut.Size           = New-Object System.Drawing.Size(1150, 720)
    $formOut.StartPosition  = "CenterScreen"
    $formOut.BackColor      = $clrBg
    $formOut.Font           = $fontNormal

    # Titel
    $pnlOutTitle = New-Object System.Windows.Forms.Panel
    $pnlOutTitle.Dock      = "Top"
    $pnlOutTitle.Height    = 56
    $pnlOutTitle.BackColor = $clrAccent
    $formOut.Controls.Add($pnlOutTitle)

    $lblOutTitle = New-Object System.Windows.Forms.Label
    $lblOutTitle.Text      = "  Event-Ergebnisse"
    $lblOutTitle.Font      = $fontTitle
    $lblOutTitle.ForeColor = [System.Drawing.Color]::White
    $lblOutTitle.Dock      = "Fill"
    $lblOutTitle.TextAlign = "MiddleLeft"
    $pnlOutTitle.Controls.Add($lblOutTitle)

    $lblOutSub = New-Object System.Windows.Forms.Label
    $lblOutSub.Text      = "   $($sorted.Count) Einträge  ·  $zeitText  ·  Computer: $computerLabel"
    $lblOutSub.Dock      = "Bottom"
    $lblOutSub.Height    = 18
    $lblOutSub.Font      = $fontSmall
    $lblOutSub.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 255)
    $lblOutSub.BackColor = $clrAccent
    $pnlOutTitle.Controls.Add($lblOutSub)

    # ── Filter-Zeile ──────────────────────────────────────────
    $pnlFilter = New-Object System.Windows.Forms.Panel
    $pnlFilter.Location  = New-Object System.Drawing.Point(10, 64)
    $pnlFilter.Size      = New-Object System.Drawing.Size(1115, 72)
    $pnlFilter.BackColor = $clrPanel
    $pnlFilter.BorderStyle = "None"
    $formOut.Controls.Add($pnlFilter)

    $pnlFilter.Add_Paint({
        param($s,$e)
        $pen = New-Object System.Drawing.Pen($clrBorder, 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width-1, $s.Height-1)
    })

    $lblFil = New-Label "Suche:" 8 12 55 20 $true
    $pnlFilter.Controls.Add($lblFil)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location    = New-Object System.Drawing.Point(68, 10)
    $txtFilter.Size        = New-Object System.Drawing.Size(280, 22)
    $txtFilter.Font        = $fontNormal
    $txtFilter.ForeColor   = $clrMuted
    $txtFilter.Text        = "Suche in ID, Quelle, Nachricht..."
    $pnlFilter.Controls.Add($txtFilter)

    $txtFilter.Add_Enter({
        if ($txtFilter.Text -eq "Suche in ID, Quelle, Nachricht...") {
            $txtFilter.Text = ""
            $txtFilter.ForeColor = $clrText
        }
    })
    $txtFilter.Add_Leave({
        if ($txtFilter.Text -eq "") {
            $txtFilter.Text = "Suche in ID, Quelle, Nachricht..."
            $txtFilter.ForeColor = $clrMuted
        }
    })

    # Typ-Filter
    $pnlFilter.Controls.Add((New-Label "Typ:" 425 12 35 20))
    $cbTypFilter = New-Object System.Windows.Forms.ComboBox
    $cbTypFilter.Location      = New-Object System.Drawing.Point(465, 10)
    $cbTypFilter.Size          = New-Object System.Drawing.Size(120, 22)
    $cbTypFilter.DropDownStyle = "DropDownList"
    $cbTypFilter.Font          = $fontNormal
    @("Alle","Critical","Error","Warning","Information","Verbose") |
        ForEach-Object { $cbTypFilter.Items.Add($_) | Out-Null }
    $cbTypFilter.SelectedIndex = 0
    $pnlFilter.Controls.Add($cbTypFilter)

    # Protokoll-Filter (dynamisch aus den Ergebnissen)
    $pnlFilter.Controls.Add((New-Label "Protokoll:" 565 12 70 20))
    $cbLogFilter = New-Object System.Windows.Forms.ComboBox
    $cbLogFilter.Location      = New-Object System.Drawing.Point(638, 10)
    $cbLogFilter.Size          = New-Object System.Drawing.Size(105, 22)
    $cbLogFilter.DropDownStyle = "DropDownList"
    $cbLogFilter.Font          = $fontNormal
    $cbLogFilter.Items.Add("Alle") | Out-Null
    $sorted | Select-Object -ExpandProperty Protokoll -Unique | Sort-Object |
        ForEach-Object { $cbLogFilter.Items.Add($_) | Out-Null }
    $cbLogFilter.SelectedIndex = 0
    $pnlFilter.Controls.Add($cbLogFilter)

    # Computer-Filter (dynamisch aus den Ergebnissen)
    $pnlFilter.Controls.Add((New-Label "Computer:" 750 12 68 20))
    $cbCompFilter = New-Object System.Windows.Forms.ComboBox
    $cbCompFilter.Location      = New-Object System.Drawing.Point(820, 10)
    $cbCompFilter.Size          = New-Object System.Drawing.Size(130, 22)
    $cbCompFilter.DropDownStyle = "DropDownList"
    $cbCompFilter.Font          = $fontNormal
    $cbCompFilter.Items.Add("Alle") | Out-Null
    $sorted | Select-Object -ExpandProperty Computer -Unique | Sort-Object |
        ForEach-Object { $cbCompFilter.Items.Add($_) | Out-Null }
    $cbCompFilter.SelectedIndex = 0
    $pnlFilter.Controls.Add($cbCompFilter)

    # Ergebnis-Zähler
    $lblCount = New-Label "" 960 12 145 20 $false $true
    $pnlFilter.Controls.Add($lblCount)

    # ── Zeile 2: Aktions-Buttons ─────────────────────────────────
    $btnExport      = New-StyledButton "CSV-Export"   8   44 100 24 $false
    $btnExportExcel = New-StyledButton "Excel-Export" 114  44 105 24 $false
    $btnChart       = New-StyledButton "Diagramm"    225  44 100 24 $false
    $pnlFilter.Controls.Add($btnExport)
    $pnlFilter.Controls.Add($btnExportExcel)
    $pnlFilter.Controls.Add($btnChart)

    # ── DataGridView ──────────────────────────────────────────
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location              = New-Object System.Drawing.Point(10, 142)
    $dgv.Size                  = New-Object System.Drawing.Size(1115, 472)
    $dgv.BackgroundColor       = $clrPanel
    $dgv.GridColor             = $clrBorder
    $dgv.BorderStyle           = "None"
    $dgv.Font                  = $fontSmall
    $dgv.DefaultCellStyle.Font = $fontSmall
    $dgv.DefaultCellStyle.ForeColor    = $clrText
    $dgv.DefaultCellStyle.BackColor    = $clrPanel
    $dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(220,225,255)
    $dgv.DefaultCellStyle.SelectionForeColor = $clrText
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235,235,248)
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = $clrAccent
    $dgv.ColumnHeadersDefaultCellStyle.Font      = $fontBold
    $dgv.ColumnHeadersHeight      = 28
    $dgv.RowHeadersVisible        = $false
    $dgv.AllowUserToAddRows       = $false
    $dgv.AllowUserToDeleteRows    = $false
    $dgv.ReadOnly                 = $true
    $dgv.SelectionMode            = "FullRowSelect"
    $dgv.AutoSizeRowsMode         = "None"
    $dgv.RowTemplate.Height       = 22
    $dgv.EnableHeadersVisualStyles = $false
    $dgv.ScrollBars               = "Both"
    $formOut.Controls.Add($dgv)

    # Spalten definieren
    $cols = @(
        @{Name="Computer";     Width=100; Fill=$false}
        @{Name="Zeit";         Width=135; Fill=$false}
        @{Name="Typ";          Width=80;  Fill=$false}
        @{Name="Protokoll";    Width=85;  Fill=$false}
        @{Name="EventID";      Width=65;  Fill=$false}
        @{Name="Kategorie";    Width=90;  Fill=$false}
        @{Name="Quelle";       Width=150; Fill=$false}
        @{Name="Beschreibung"; Width=185; Fill=$false}
        @{Name="Nachricht";    Width=0;   Fill=$true}
    )
    foreach ($col in $cols) {
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.Name        = $col.Name
        $c.HeaderText  = $col.Name
        $c.SortMode    = "Automatic"
        if ($col.Fill) {
            $c.AutoSizeMode = "Fill"
        } else {
            $c.Width = $col.Width
        }
        $dgv.Columns.Add($c) | Out-Null
    }

    # Zeilenfärbung nach Typ
    $dgv.Add_CellFormatting({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $row = $s.Rows[$e.RowIndex]
        $typ = $row.Cells["Typ"].Value
        $e.CellStyle.BackColor = switch ($typ) {
            "Critical"    { [System.Drawing.Color]::FromArgb(255, 225, 225) }
            "Error"       { [System.Drawing.Color]::FromArgb(255, 240, 240) }
            "Warning"     { [System.Drawing.Color]::FromArgb(255, 252, 235) }
            "Information" { $clrPanel }
            "Verbose"     { [System.Drawing.Color]::FromArgb(245, 245, 250) }
            default       { $clrPanel }
        }
    })

    # Funktion: Tabelle befüllen
    function Update-Grid($data) {
        $dgv.Rows.Clear()
        foreach ($r in $data) {
            $dgv.Rows.Add(
                $r.Computer,
                $r.Zeit.ToString("dd.MM.yyyy HH:mm:ss"),
                $r.Typ,
                $r.Protokoll,
                $r.EventID,
                $r.Kategorie,
                $r.Quelle,
                $r.Beschr,
                $r.Nachricht
            ) | Out-Null
        }
        $lblCount.Text = "$($data.Count) / $($sorted.Count)"
    }

    # Filter-Funktion
    function Apply-Filter {
        $fText = $txtFilter.Text.Trim()
        if ($fText -eq "Suche in ID, Quelle, Nachricht...") { $fText = "" }
        $fTyp  = $cbTypFilter.SelectedItem
        $fLog  = $cbLogFilter.SelectedItem
        $fComp = $cbCompFilter.SelectedItem

        $filtered = $sorted | Where-Object {
            $okText = ($fText -eq "") -or
                      ($_.EventID -like "*$fText*") -or
                      ($_.Quelle  -like "*$fText*") -or
                      ($_.Nachricht -like "*$fText*") -or
                      ($_.Beschr  -like "*$fText*")
            $okTyp  = ($fTyp -eq "Alle") -or ($_.Typ -eq $fTyp)
            $okLog  = ($fLog -eq "Alle") -or ($_.Protokoll -eq $fLog)
            $okComp = ($fComp -eq "Alle") -or ($_.Computer -eq $fComp)
            $okText -and $okTyp -and $okLog -and $okComp
        }
        Update-Grid $filtered
    }

    $txtFilter.Add_TextChanged({ Apply-Filter })
    $cbTypFilter.Add_SelectedIndexChanged({ Apply-Filter })
    $cbLogFilter.Add_SelectedIndexChanged({ Apply-Filter })
    $cbCompFilter.Add_SelectedIndexChanged({ Apply-Filter })

    # Initial befüllen
    Update-Grid $sorted

    # Detail-Anzeige bei Klick
    $lblDetail = New-Object System.Windows.Forms.Label
    $lblDetail.Location  = New-Object System.Drawing.Point(10, 620)
    $lblDetail.Size      = New-Object System.Drawing.Size(1115, 40)
    $lblDetail.Font      = $fontSmall
    $lblDetail.ForeColor = $clrMuted
    $lblDetail.Text      = "Zeile anklicken für vollständige Nachricht"
    $formOut.Controls.Add($lblDetail)

    $dgv.Add_SelectionChanged({
        if ($dgv.SelectedRows.Count -gt 0) {
            $row = $dgv.SelectedRows[0]
            $msg = $row.Cells["Nachricht"].Value
            $lblDetail.Text = "📄 " + $msg.Substring(0, [Math]::Min(200, $msg.Length))
            $lblDetail.ForeColor = $clrText
        }
    })

    # Doppelklick → vollständige Nachricht
    $dgv.Add_CellDoubleClick({
        if ($dgv.SelectedRows.Count -gt 0) {
            $row   = $dgv.SelectedRows[0]
            $detail = "Computer:    $($row.Cells['Computer'].Value)`n" +
                      "Zeitstempel: $($row.Cells['Zeit'].Value)`n" +
                      "Typ:         $($row.Cells['Typ'].Value)`n" +
                      "Protokoll:   $($row.Cells['Protokoll'].Value)`n" +
                      "Event-ID:    $($row.Cells['EventID'].Value)`n" +
                      "Quelle:      $($row.Cells['Quelle'].Value)`n" +
                      "Kategorie:   $($row.Cells['Kategorie'].Value)`n" +
                      "Beschr.:     $($row.Cells['Beschreibung'].Value)`n`n" +
                      "--- Vollständige Nachricht ---`n" +
                      $row.Cells["Nachricht"].Value
            [System.Windows.Forms.MessageBox]::Show($detail, "Event-Detail", "OK", "Information") | Out-Null
        }
    })

    # ── CSV-Export ────────────────────────────────────────────────
    $btnExport.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter   = "CSV-Datei (*.csv)|*.csv"
        $dlg.FileName = "Events_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($dlg.ShowDialog() -eq "OK") {
            $sorted | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ";"
            [System.Windows.Forms.MessageBox]::Show(
                "CSV-Export abgeschlossen:`n$($dlg.FileName)",
                "Export erfolgreich", "OK", "Information") | Out-Null
        }
    })

    # ── Excel-Export ─────────────────────────────────────────────
    $btnExportExcel.Add_Click({
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Microsoft Excel ist nicht installiert oder nicht verfügbar.`nBitte CSV-Export verwenden.",
                "Excel nicht gefunden", "OK", "Warning") | Out-Null
            return
        }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter   = "Excel-Datei (*.xlsx)|*.xlsx"
        $dlg.FileName = "Events_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
        if ($dlg.ShowDialog() -ne "OK") {
            try { $excel.Quit() } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            return
        }
        try {
            $excel.Visible      = $false
            $excel.DisplayAlerts = $false
            $wb = $excel.Workbooks.Add()
            $ws = $wb.Worksheets.Item(1)
            $ws.Name = "Events"

            # Header
            $headers = @("Computer","Zeit","Typ","Protokoll","EventID","Kategorie","Quelle","Beschreibung","Nachricht")
            for ($col = 0; $col -lt $headers.Count; $col++) {
                $ws.Cells.Item(1, $col+1)                    = $headers[$col]
                $ws.Cells.Item(1, $col+1).Font.Bold          = $true
                $ws.Cells.Item(1, $col+1).Interior.ColorIndex = 37
            }

            # Daten
            $rowNum = 2
            foreach ($r in $sorted) {
                $ws.Cells.Item($rowNum, 1) = $r.Computer
                $ws.Cells.Item($rowNum, 2) = $r.Zeit.ToString("dd.MM.yyyy HH:mm:ss")
                $ws.Cells.Item($rowNum, 3) = $r.Typ
                $ws.Cells.Item($rowNum, 4) = $r.Protokoll
                $ws.Cells.Item($rowNum, 5) = $r.EventID
                $ws.Cells.Item($rowNum, 6) = $r.Kategorie
                $ws.Cells.Item($rowNum, 7) = $r.Quelle
                $ws.Cells.Item($rowNum, 8) = $r.Beschr
                $ws.Cells.Item($rowNum, 9) = $r.Nachricht
                # Zeilenfarbe nach Typ
                $ci = switch ($r.Typ) {
                    "Critical" { 3 }   # rot
                    "Error"    { 45 }  # lachs
                    "Warning"  { 36 }  # gelb
                    default    { 0 }   # keine
                }
                if ($ci -gt 0) {
                    $ws.Range($ws.Cells.Item($rowNum,1), $ws.Cells.Item($rowNum,9)).Interior.ColorIndex = $ci
                }
                $rowNum++
            }

            $ws.Columns.AutoFit() | Out-Null
            # Nachricht-Spalte auf max. 80 Zeichen
            if ($ws.Columns.Item(9).ColumnWidth -gt 80) { $ws.Columns.Item(9).ColumnWidth = 80 }

            $wb.SaveAs($dlg.FileName, 51)  # 51 = xlOpenXMLWorkbook
            $wb.Close($false)
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)    | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)    | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

            [System.Windows.Forms.MessageBox]::Show(
                "Excel-Export abgeschlossen:`n$($dlg.FileName)",
                "Export erfolgreich", "OK", "Information") | Out-Null
        } catch {
            try { $excel.Quit() } catch {}
            [System.Windows.Forms.MessageBox]::Show(
                "Excel-Export fehlgeschlagen:`n$($_.Exception.Message)",
                "Fehler", "OK", "Error") | Out-Null
        }
    })

    # ── Zeitreihe-Diagramm ───────────────────────────────────────
    $btnChart.Add_Click({
        # Zeitspanne bestimmen -> Granularität Stunde oder Tag
        $useDay = $sorted.Count -gt 0 -and ($sorted[0].Zeit - $sorted[-1].Zeit).TotalHours -gt 48
        $grouped = @($sorted | Group-Object {
            if ($useDay) { $_.Zeit.ToString("dd.MM.yyyy") }
            else         { $_.Zeit.ToString("dd.MM. HH:00") }
        })
        [Array]::Reverse($grouped)   # aelteste links, neueste rechts
        $labels = @($grouped | ForEach-Object { $_.Name })
        $values = @($grouped | ForEach-Object { $_.Count })

        $fChart = New-Object System.Windows.Forms.Form
        $fChart.Text          = "Zeitreihe - Event-Häufigkeit"
        $fChart.Size          = New-Object System.Drawing.Size(950, 520)
        $fChart.StartPosition = "CenterScreen"
        $fChart.BackColor     = $clrBg
        $fChart.Font          = $fontNormal

        try {
            $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
            $chart.Dock      = "Fill"
            $chart.BackColor = $clrBg

            $ca = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
            $ca.BackColor          = [System.Drawing.Color]::White
            $ca.AxisX.Interval     = [Math]::Max(1, [Math]::Ceiling($labels.Count / 20))
            $ca.AxisX.LabelStyle.Angle = -45
            $ca.AxisX.LabelStyle.Font  = $fontSmall
            $ca.AxisY.Title        = "Anzahl Events"
            $ca.AxisY.TitleFont    = $fontSmall
            $ca.AxisY.LabelStyle.Font  = $fontSmall
            $chart.ChartAreas.Add($ca)

            $title = New-Object System.Windows.Forms.DataVisualization.Charting.Title
            $title.Text = "Event-Häufigkeit  ·  $($sorted.Count) Events  ·  $computerLabel"
            $title.Font = $fontBold
            $chart.Titles.Add($title)

            $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
            $series.ChartType   = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
            $series.Color       = $clrAccent
            $series.BorderColor = [System.Drawing.Color]::White
            $series.BorderWidth = 1
            for ($i = 0; $i -lt $labels.Count; $i++) {
                $series.Points.AddXY($labels[$i], $values[$i]) | Out-Null
            }
            $chart.Series.Add($series)

            $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
            $legend.Enabled = $false
            $chart.Legends.Add($legend)

            $fChart.Controls.Add($chart)
        } catch {
            $errLbl = New-Label "⚠  Diagramm-Assembly nicht verfügbar: $($_.Exception.Message)" 20 20 860 40
            $errLbl.ForeColor = $clrWarn
            $fChart.Controls.Add($errLbl)
        }
        $fChart.ShowDialog() | Out-Null
    })

    # Fehler-Hinweis
    if ($errors.Count -gt 0) {
        $lblErr = New-Label "⚠  $($errors.Count) Protokoll(e) nicht lesbar (Adminrechte prüfen)" 10 665 700 20 $false
        $lblErr.ForeColor = $clrWarn
        $formOut.Controls.Add($lblErr)
    }

    # ── Live-Modus ────────────────────────────────────────────
    $lblLiveStatus = New-Object System.Windows.Forms.Label
    $lblLiveStatus.Location  = New-Object System.Drawing.Point(10, 668)
    $lblLiveStatus.Size      = New-Object System.Drawing.Size(500, 18)
    $lblLiveStatus.Font      = $fontSmall
    $lblLiveStatus.ForeColor = $clrMuted
    $formOut.Controls.Add($lblLiveStatus)

    if ($chkLive.Checked) {
        # Interval bestimmen
        $script:liveTimer.Interval = switch ($cbLiveInterval.SelectedItem) {
            "30 Sek"      { 30000  }
            "1 Minute"    { 60000  }
            "5 Minuten"   { 300000 }
            "10 Minuten"  { 600000 }
            default       { 60000  }
        }

        # Abfrage-Parameter für Timer einfrieren (Closure über $checkedIDs, $startTime, $maxCount, $computer, $cred)
        $script:liveTimer.Add_Tick({
            $freshResults = [System.Collections.Generic.List[PSObject]]::new()
            $byLogLive = $checkedIDs | Group-Object -Property Log
            foreach ($grp in $byLogLive) {
                $logName = $grp.Name
                $allIds  = @($grp.Group | Select-Object -ExpandProperty ID -Unique)
                $idLookup = @{}
                foreach ($ev in $grp.Group) { $idLookup[$ev.ID] = $ev }
                for ($offset = 0; $offset -lt $allIds.Count; $offset += 22) {
                    $end      = [Math]::Min($offset + 21, $allIds.Count - 1)
                    $idsBatch = $allIds[$offset..$end]
                    $filter   = @{ LogName = $logName; ID = $idsBatch }
                    if ($startTime) { $filter.StartTime = $startTime }
                    try {
                        $qp = @{ ComputerName = $computer; FilterHashtable = $filter; MaxEvents = $maxCount; ErrorAction = 'Stop' }
                        if ($cred) { $qp.Credential = $cred }
                        $winEvents = Get-WinEvent @qp
                        foreach ($r in $winEvents) {
                            $meta  = $idLookup[[int]$r.Id]
                            $msg   = if ($r.Message) { $r.Message } else { "(keine Nachricht)" }
                            $short = ($msg -replace "`r`n|`n", " ")
                            if ($short.Length -gt 300) { $short = $short.Substring(0, 300) + "..." }
                            $typ = switch ($r.LevelDisplayName) {
                                "Fehler"        { "Error" }     "Error"         { "Error" }
                                "Kritisch"      { "Critical" }  "Critical"      { "Critical" }
                                "Warnung"       { "Warning" }   "Warning"       { "Warning" }
                                "Information"   { "Information" } "Informationen" { "Information" }
                                default         { "$($r.LevelDisplayName)" }
                            }
                            $freshResults.Add([PSCustomObject]@{
                                Zeit = $r.TimeCreated; Typ = $typ
                                Protokoll = Get-LogShortLabel $logName; LogVoll = $logName
                                EventID = $r.Id; Quelle = $r.ProviderName
                                Kategorie = if ($meta) { $meta.Category } else { "" }
                                Beschr    = if ($meta) { $meta.Desc }     else { "" }
                                Nachricht = $short
                            })
                        }
                    } catch {}
                }
            }
            $script:liveDataRef = $freshResults | Sort-Object Zeit -Descending
            Update-Grid $script:liveDataRef
            $lblLiveStatus.ForeColor = $clrSuccess
            $lblLiveStatus.Text = "Live: letzte Aktualisierung $(Get-Date -Format 'HH:mm:ss')  –  $($script:liveDataRef.Count) Einträge"
        })

        $script:liveTimer.Start()
        $lblLiveStatus.ForeColor = $clrInfo
        $lblLiveStatus.Text = "Live-Modus aktiv – Intervall: $($cbLiveInterval.SelectedItem)"

        $formOut.Add_FormClosed({
            $script:liveTimer.Stop()
            $script:liveTimer.Remove_Tick($script:liveTimer.Tag)
            $lblLiveStatus.Text = ""
        })
    }

    $formOut.ShowDialog() | Out-Null
    $script:liveTimer.Stop()
})

# ── Starten ───────────────────────────────────────────────────
[System.Windows.Forms.Application]::Run($formMain)