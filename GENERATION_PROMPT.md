# Generation Prompt – Windows Event Analyzer

> Dieser Prompt beschreibt das Skript vollständig und präzise genug,  
> um es neu zu erzeugen. Sprache: Deutsch/English gemischt (wie Original).

---

## Aufgabe

Erstelle ein **Windows PowerShell 5.1-kompatibles GUI-Skript** (`WindowsEventAnalyzer.ps1`),
das IT-Administratoren ermöglicht, Windows-Ereignisprotokolle (Event Logs) lokal und
remote interaktiv abzufragen, zu filtern und zu exportieren.
Das Skript läuft als eigenständige WinForms-Anwendung, benötigt **keine externen Abhängigkeiten**
außer .NET Framework (System.Windows.Forms, System.Drawing) und kann mit ps2exe zu einer EXE
kompiliert werden.

---

## Anforderungen im Detail

### 1. Auto-Elevation

Das Skript muss sich beim Start **automatisch als Administrator neu starten**, falls es ohne
Adminrechte ausgeführt wird:

```powershell
$_principal = New-Object Security.Principal.WindowsPrincipal(
    [Security.Principal.WindowsIdentity]::GetCurrent()
)
if (-not $_principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
Remove-Variable _principal
```

### 2. Rechtlicher Hinweis (Disclaimer)

Am Anfang des Skripts steht ein mehrzeiliger Kommentar-Block mit:
- Nutzungszweck (nur autorisierte IT-Administratoren)
- Hinweis auf Strafbarkeit bei unbefugtem Zugriff (§202a StGB, §118a öStGB, Art. 143 CH-StGB,
  Richtlinie 2013/40/EU, CFAA USA)
- DSGVO-Hinweis (Ereignisprotokolle können personenbezogene Daten enthalten)
- Credentials-Hinweis (SecureString, keine Persistenz)
- Haftungsausschluss
- ExecutionPolicy-Empfehlung

### 3. Design / Farbschema

Einheitliches, modernes Flat-Design mit diesen Farben:

| Variable       | RGB                | Verwendung                          |
|----------------|--------------------|-------------------------------------|
| `$clrBg`       | 245, 245, 250      | Fensterhintergrund                  |
| `$clrPanel`    | 255, 255, 255      | Panels / Karten                     |
| `$clrAccent`   | 74, 74, 170        | Titelleisten, primäre Buttons       |
| `$clrAccentHov`| 94, 94, 190        | Hover-Farbe für Buttons             |
| `$clrBorder`   | 210, 210, 230      | Panel-Rahmen                        |
| `$clrText`     | 30, 30, 50         | Normaler Text                       |
| `$clrMuted`    | 110, 110, 140      | Hinweistexte / Platzhalter          |
| `$clrWarn`     | 200, 80, 60        | Fehler/Warnungen                    |
| `$clrSuccess`  | 40, 160, 80        | Erfolgsmeldungen                    |
| `$clrInfo`     | 30, 120, 200       | Infomeldungen                       |

Fonts: Segoe UI 9 (normal/bold/small), Segoe UI 12 Bold (Titel), Consolas 8 (Mono für Listen).

Hilfsfunktionen:
- `New-StyledButton($text, $x, $y, $w, $h, $primary)` – primary=blau, secondary=weiß mit Rand
- `New-Label($text, $x, $y, $w, $h, $bold, $muted)`
- `New-SectionPanel($x, $y, $w, $h, $title)` – weißes Panel mit Paint-Event-Rand und Titelzeile

### 4. Event-ID Katalog

Statischer Katalog als `[System.Collections.Generic.List[PSObject]]` mit Feldern:
`ID` (int), `Log` (vollständiger Log-Name), `Category` (string), `Icon` (●▲○◆◈), `Desc` (string).

Enthält kuratierte Events für folgende Log-Gruppen:

| Gruppe             | Log-Name (Beispiel)                                          | Typische IDs                          |
|--------------------|--------------------------------------------------------------|---------------------------------------|
| System             | `System`                                                     | 41, 55, 11, 1074, 6005/6006/6008, 7000-7045 |
| Security           | `Security`                                                   | 4624/4625, 4647/4648, 4688, 4697-4702, 4720-4740, 4768-4776, 4946-4948, 5156/5157 |
| Application        | `Application`                                                | 1000, 1001, 1002, 1026, 11707/11708/11724 |
| PowerShell         | `Microsoft-Windows-PowerShell/Operational`                   | 4100, 4103, 4104, 4105, 4106, 40961/40962 |
| Windows Defender   | `Microsoft-Windows-Windows Defender/Operational`             | 1006-1009, 1015, 1116/1117, 2000/2001, 5001/5007/5010 |
| WLAN               | `Microsoft-Windows-WLAN-AutoConfig/Operational`              | 8001-8003, 11001/11004/11006 |
| Task Scheduler     | `Microsoft-Windows-TaskScheduler/Operational`                | 106, 129, 140, 141, 200-203 |
| RDP (LSM)          | `Microsoft-Windows-TerminalServices-LocalSessionManager/Operational` | 21-25, 39, 40 |
| RDP (RCM)          | `Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational` | 1149, 1150, 1158 |
| Druckerdienst      | `Microsoft-Windows-PrintService/Operational`                 | 307, 805, 842 |
| BitLocker          | `Microsoft-Windows-BitLocker/BitLocker Management`           | 845, 846, 24620 |
| SMB Client         | `Microsoft-Windows-SmbClient/Connectivity`                   | 30803, 30806 |

### 5. Computer-Scan (`Invoke-ComputerEventScan`)

Funktion, die **alle aktiven Event-Logs** eines Computers durchsucht und vorhandene Event-IDs erkennt:

```
Parameter: Computer, ProgressUI, MaxPerLog=15, MaxLogs=120, Credential=$null
```

- Unterscheidet lokal vs. remote via `Test-IsLocalComputer` (prüft Hostname, localhost, IP, FQDN)
- Lokal: kein `ComputerName`-Parameter (WinRM-unabhängig)
- Remote: `ComputerName` + optional `Credential`
- **Schritt 1:** `Get-WinEvent -ListLog '*' -ErrorAction SilentlyContinue`
  filtert auf `IsEnabled -and RecordCount > 0`, sortiert nach RecordCount absteigend, Top-N
- **Schritt 2:** Pro Log ein Sample lesen (MaxPerLog Events), nach ID gruppieren, ohne `.Message`-Zugriff
  (kein DLL-Aufruf – wichtig für Geschwindigkeit)
- Beschreibung im Scan-Ergebnis: `"[Provider: $provider · $count x in Sample · Log: $logName]"`
- Fehlermeldungen harmloser Art (Log leer, nicht vorhanden) werden ignoriert
- Abbruch-Unterstützung: `$script:scanCancelled` (via "Überspringen"-Button)
- Progress-Fenster (`New-ProgressForm`): blaue Titelleiste, Fortschrittsbalken, Sub-Label, TopMost=true

**`Merge-DiscoveredEvents`**: Fügt gefundene IDs in den globalen Katalog ein (Deduplizierung via "ID|Log"-Key),
merkt erkannte Log-Namen in `$script:discoveredLogs`.

**Manifest-Scan** (`Invoke-ManifestEventScan`): Liest `Get-WinEvent -ListProvider '*'` und
extrahiert alle dokumentierten Event-IDs aus Provider-Manifesten als Kategorie "Dokumentiert".

### 6. Startup-Verhalten

Beim Start:
1. Temporäres TopMost-Fenster (1×1 px, unsichtbar) als Owner für MessageBox
2. MessageBox: „Lokalen Computer '<HOSTNAME>' beim Start scannen?"  [Ja/Nein]
3. Bei Ja: Progress-Form anzeigen, `Invoke-ComputerEventScan` ausführen
4. Bei 0 Treffern: `Show-ScanDiagnostics` anzeigen (Diagnose-Dialog)
5. Ergebnis in `$script:startupScanResult` speichern (wird in Statusbar angezeigt)

### 7. Formular 1 – Auswahl-Maske

**Fenstergröße:** 920 × 862 px, FixedSingle, kein Maximize.
**TopMost = $true** nur initial; nach `Add_Shown` → `TopMost = $false; Activate()`.

#### Sektion „Abfrage-Optionen" (SectionPanel, y=75):

| Control         | Typ         | Inhalt / Funktion                                     |
|-----------------|-------------|-------------------------------------------------------|
| Zeitraum        | ComboBox    | Letzte 1h / 6h / 24h / 7 Tage / 30 Tage / Alles; Default: 24h |
| Max. Einträge   | ComboBox    | 25 / 50 / 100 / 250 / 500; Default: 50               |
| Computer        | TextBox     | Hostname oder IP, kommagetrennt für Mehrfach-Computer; Default: `$env:COMPUTERNAME` |
| Scan-Button     | Button      | Führt `Invoke-ComputerEventScan` erneut aus           |
| Manifest-Checkbox | CheckBox  | Aktiviert zusätzlichen Manifest-Scan                  |
| Domain          | TextBox     | AD-Domäne für Remote-Credentials                      |
| Benutzer        | TextBox     | Benutzername für Remote                               |
| Passwort        | TextBox     | PasswordChar='*', ConvertTo-SecureString, nicht persistiert |
| Profil          | ComboBox    | Dropdown gespeicherter Profile (JSON-Dateien)         |
| Laden/Speichern | Buttons     | Profile laden/speichern                               |
| Live-Modus      | CheckBox    | Aktiviert automatische Aktualisierung                 |
| Live-Intervall  | ComboBox    | 30 Sek / 1 Minute / 5 Minuten / 10 Minuten           |

Alle Controls haben **ToolTips** (AutoPopDelay 9000ms, InitialDelay 450ms).

#### Sektion „Events auswählen" (SectionPanel, y=245):

- **Kategorie-Filter** ComboBox: Alle / Kritisch / Fehler / Warnung / Information / Überwachung / Eigene / Erkannt / Dokumentiert
- **Bereich-Filter** ComboBox: Alle / >> Gefunden / ★ Empfohlen / ◈ Erkannt / = Dokumentiert / System / Security / Application / PowerShell / Defender / WLAN / TaskScheduler / Remote Desktop / Druckerdienst / BitLocker / SMB / Eigene IDs
- **[+] Alle / [-] Keine** Buttons (nur auf sichtbare Einträge)
- **Volltextsuche** TextBox (filtert sofort nach ID, Desc, Category, Log)
- **CheckedListBox** (Consolas 9pt, CheckOnClick, HorizontalScrollbar):
  Jede Zeile: `Icon  ID   [LOG-LABEL]  Kategorie  –  Beschreibung`
- `$script:visibleEvents` List: Mapping CLB-Index → Katalog-Eintrag
- **Eigene Event-ID** Eingabe: TextBox + Button „Hinzufügen" (Format: `ID@LogName` oder nur `ID` für System)

#### Sektion „Aktionen" (y=625):

- **[Abfragen]** Button (primär, blau)
- **[Test: Lokal]** Button – schneller Test nur lokaler Computer, alle gewählten IDs
- **[Test: Remote]** Button – Test mit eingetragenem Computer/Credentials
- **[Diagnose]** Button – `Show-ScanDiagnostics` Dialog
- **Fortschrittsbalken** (während Abfrage, sonst unsichtbar)
- **Status-Label** (Erfolg/Fehler/Hinweise, farbkodiert)
- Statuszeile unten: Scan-Ergebnis vom Startup-Scan

### 8. Abfrage-Engine (`Invoke-Query`)

```
Parameter: Computers (string[]), SelectedEvents (PSObject[]), MaxCount, StartTime, Credential
```

- Für jeden Computer:
  - `Test-IsLocalComputer` → lokal oder remote
  - Events nach Log-Name gruppieren
  - Pro Log: IDs in 22er-Batches aufteilen (Get-WinEvent -FilterHashtable unterstützt max. 22 IDs!)
  - `Get-WinEvent -FilterHashtable @{ LogName=...; ID=...; StartTime=... } -MaxEvents $maxCount`
  - Credentials nur bei Remote
- Pro Ergebnis-Zeile:
  - `Computer`, `Zeit` (TimeCreated), `Typ` (normalisiert: Error/Critical/Warning/Information/Verbose)
  - `Protokoll` (Kurz-Label via `Get-LogShortLabel`), `LogVoll`, `EventID`, `Quelle` (ProviderName)
  - `Kategorie` (aus Katalog), `Beschr` (s.u.), `Nachricht` ($r.Message, Zeilenumbrüche → Leerzeichen, **keine Längen-Begrenzung**)
- **Beschreibung-Logik:**
  Wenn `$meta.Desc` vorhanden und nicht wie `^\[Provider:` aussieht → Katalog-Text verwenden.
  Sonst: erste nicht-leere Zeile der Nachricht, max. 120 Zeichen + „…".
- Harmlose Fehler (Log leer/nicht vorhanden) werden ignoriert; echte Fehler werden gesammelt
- Bei 0 Treffern: Diagnose-Dialog mit Abfrage-Statistik
- Ergebnisse nach Computer + Zeit absteigend sortiert

### 9. Formular 2 – Ergebnis-Ausgabe

**Fenstergröße:** 1150 × 720 px, resizable, TopMost = $true.

#### Filter-Zeile (Panel oben):

- Volltextsuche (filtert in ID, Quelle, Nachricht, Beschreibung)
- Typ-Filter ComboBox: Alle / Critical / Error / Warning / Information / Verbose
- Protokoll-Filter ComboBox (dynamisch aus Ergebnissen befüllt)
- Computer-Filter ComboBox (dynamisch aus Ergebnissen befüllt)
- Ergebnis-Zähler Label: „N / Gesamt"

#### Aktions-Buttons (zweite Zeile im Filter-Panel):

- **CSV-Export**: SaveFileDialog, `Export-Csv -Delimiter ";"`, UTF8
- **Excel-Export** (nur wenn Excel COM-Objekt verfügbar): native COM-Automation, farbige Zeilen
  (Critical=rot, Error=lachs, Warning=gelb), AutoFit, Nachricht-Spalte max. 80 Zeichenbreite
- **Diagramm** (Zeitreihe): falls `System.Windows.Forms.DataVisualization` verfügbar,
  sonst Fallback auf Textausgabe. X-Achse: Zeit (Stunden oder Tage je nach Zeitraum),
  Y-Achse: Anzahl, pro Typ eine Serie (rot/orange/gelb/grau)

#### DataGridView:

Spalten: Computer (100) | Zeit (135) | Typ (80) | Protokoll (85) | EventID (65) |
         Kategorie (90) | Quelle (150) | Beschreibung (185) | Nachricht (Fill)

- RowHeight = 22, kein RowHeader, ReadOnly, FullRowSelect
- **Zeilenfärbung** via CellFormatting: Critical=helles rot, Error=sehr helles rot,
  Warning=helles gelb, Info=weiß, Verbose=grau-weiß
- **Einfachklick**: erste 200 Zeichen der Nachricht im Detail-Label unter dem Grid anzeigen
- **Doppelklick**: vollständige Event-Details in Show-CopyableDialog (kopierbar)

### 10. Live-Modus

- `$script:liveTimer` (WinForms Timer), Interval aus ComboBox (30s/1min/5min/10min)
- Bei jedem Tick: für alle gewählten Logs/IDs `Get-WinEvent` ausführen, Grid aktualisieren
- Status-Label zeigt letzte Aktualisierungszeit
- **Beschreibung-Logik** (identisch mit Hauptabfrage): kein `[Provider:...]`-Text
- Timer.Stop() bei Schließen von Formular 2

### 11. Profil-System

- Speicherort: `$env:APPDATA\WindowsEventAnalyzer\profiles\<Name>.json`
- Inhalt: Name, Zeitraum-Index, MaxCount-Index, Events-Array (je `{ID, Log}`)
- `Save-Profile`: speichert geprüfte Events; warnt bei leerem Namen oder Auswahl
- `Load-Profile`: stellt Zeitraum, MaxCount und Event-Auswahl wieder her; zeigt Trefferanzahl

### 12. Show-CopyableDialog

Wiederverwendbarer Dialog mit:
- Blauer Titelleiste (Segoe UI 10 Bold, weiß)
- ReadOnly Monospace-TextBox (Consolas 9)
- „📋 Alles kopieren" Button (Toggle zu „✔ Kopiert!" mit grünem Hintergrund)
- „Schließen" Button
- TopMost = $true, Sizable (min 480×320)

### 13. Diagnose-Dialog (`Show-ScanDiagnostics`)

Zeigt bei leerem Scan-Ergebnis:
- Ob WinRM läuft (Get-Service winrm)
- Ob Firewall-Regeln für Event Log Remote Management vorhanden
- Ob der Computer erreichbar ist (Test-Connection)
- Empfehlung: WinRM aktivieren via `winrm quickconfig`

### 14. Technische Details / Fallstricke

| Problem                          | Lösung                                                          |
|----------------------------------|-----------------------------------------------------------------|
| Scan bricht bei defekten Logs ab | `Get-WinEvent -ListLog '*' -ErrorAction SilentlyContinue`       |
| Max 22 IDs pro FilterHashtable   | IDs in 22er-Batches aufteilen                                   |
| `.Message` verlangsamt Scan      | Im Scan-Modus `.Message` NICHT aufrufen                         |
| MessageBox nicht im Vordergrund  | Temporäres TopMost-Form (1×1px) als Owner übergeben             |
| PS1 ohne Admin startet ohne Rechte | Auto-Elevation via `Start-Process ... -Verb RunAs`             |
| RecordCount null bei Remote      | `$_.RecordCount -eq $null -or $_.RecordCount -gt 0`             |
| Multi-line type cast             | `New-Object Type(arg)` statt `[Type] arg` über Zeilenumbruch    |
| Profil-Verzeichnis               | Beim Start anlegen falls nicht vorhanden                        |

### 15. Kompilierung zu EXE

Das Skript ist kompatibel mit `ps2exe` (Invoke-PS2EXE):
```powershell
Invoke-PS2EXE WindowsEventAnalyzer.ps1 WindowsEventAnalyzer.exe `
    -noConsole -requireAdmin `
    -title "Windows Event Analyzer" `
    -description "Interaktives GUI-Tool..." `
    -company "FrankBKW" -product "Windows Event Analyzer" `
    -version "1.2.x.0" -copyright "(c) 2025 FrankBKW"
```
`-requireAdmin` ersetzt die Auto-Elevation im PS1 nicht – beide sind vorhanden.

---

## Nicht-Anforderungen (bewusst weggelassen)

- Kein Netzwerk-Zugriff außer WinRM zu Ziel-Computern
- Keine Datei-Schreiboperationen außer Profil-JSON und Export (CSV/XLSX)
- Keine Registry-Schreibzugriffe
- Keine persistierten Credentials
- Keine scheduled Tasks, keine Autostart-Einträge
- Keine externen PowerShell-Module zur Laufzeit
