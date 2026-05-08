# Windows Events Viewer (PowerShell GUI)

Interaktives PowerShell-WinForms-Tool zum Abfragen und Filtern von Windows-Eventlogs. Ein
Auswahlformular mit kuratiertem Event-Katalog plus automatischem Computer-Scan, ein
Ausgabeformular mit DataGridView und Live-Filter.

## Datei

- `WindowsEvents-Viewer.ps1` – komplettes Skript, ~1500 Zeilen, single-file
- Start: `PowerShell -ExecutionPolicy Bypass -File WindowsEvents-Viewer.ps1`
- Adminrechte empfohlen (Security-Log und Operational-Logs)

## Architektur

Aufbau von oben nach unten innerhalb der Datei:

1. **Add-Type / Farben / Fonts** – WinForms- und Drawing-Assemblies, Farb-Variablen (`$clr*`), Fonts
2. **Event-Katalog** – ~110 kuratierte Einträge in `[List[PSObject]]` mit Feldern `ID`, `Log`, `Category`, `Icon`, `Desc`. Aufgeteilt in Sektionen: System, Security, Application, PowerShell, Defender, WLAN, TaskScheduler, RDP-LSM, RDP-RCM, Print, BitLocker, SMB
3. **Scan-Funktionen**
   - `New-ProgressForm` – modaler Fortschrittsdialog mit Cancel-Button
   - `Invoke-ComputerEventScan` – Phase 1: Sample der letzten 500 Events pro Log → Kategorie `Erkannt`
   - `Invoke-DocumentedIDsScan` – Phase 2: Provider-Manifeste → Kategorie `Dokumentiert`
   - `Merge-DiscoveredEvents` – Dedup nach `ID|Log`-Schlüssel
4. **Helper-Funktionen** – `New-StyledButton`, `New-Label`, `New-SectionPanel`, `Get-LogShortLabel`
5. **Startup-Scan** – läuft beim Skriptstart, befüllt Katalog mit Erkannt + Dokumentiert
6. **Formular 1 (Auswahl)** – CheckedListBox + Filter-Dropdowns + Eigene-IDs-Sektion + Re-Scan
7. **Update-EventList** – synchronisiert CLB mit `$script:visibleEvents` (Index → Katalog-Eintrag)
8. **Abfrage-Logik** – Click-Handler für „Abfragen"-Button
9. **Formular 2 (Ausgabe)** – DataGridView mit Live-Filter, Farbkodierung, CSV-Export

## Wichtige Implementierungsdetails

### Get-WinEvent statt Get-EventLog
`Get-EventLog` kann nur die drei klassischen Logs lesen. `Get-WinEvent` ist
zwingend für moderne Logs (`Microsoft-Windows-...`). NIE wieder zu `Get-EventLog` zurück.

### 22-IDs-Limit von FilterHashtable
**Hartes Limit**: `Get-WinEvent -FilterHashtable @{ID=@(...)}` akzeptiert max. 22 IDs.
Wir batchen daher in der Abfrage in Gruppen von 22:
```powershell
for ($offset = 0; $offset -lt $allIds.Count; $offset += 22) {
    $end = [Math]::Min($offset + 21, $allIds.Count - 1)
    $idsBatch = $allIds[$offset..$end]
    # Get-WinEvent mit $idsBatch
}
```
Dieses Limit gilt auch für andere Hashtable-Parameter — Vorsicht beim Erweitern.

### Index-Mapping CheckedListBox <-> Katalog
**Niemals String-Parsing zur Rück-Identifikation der Katalog-Einträge!**
Die parallele Liste `$script:visibleEvents` wird in `Update-EventList` synchron
mit der CLB befüllt. Index-Position in der CLB = Index in `$script:visibleEvents`.
Frühere Versuche mit Regex/`-like` auf dem Display-String waren fehleranfällig
(Padding mit Leerzeichen brach das Match).

### Icons müssen BMP-Unicode sein
Consolas-Schrift (für Spaltenausrichtung in der CLB) hat **keine farbigen Emojis**.
Verwendet werden monochrome BMP-Zeichen:
- `●` (U+25CF) Kritisch/Fehler
- `▲` (U+25B2) Warnung
- `○` (U+25CB) Information
- `◆` (U+25C6) Überwachung
- `◈` (U+25C8) Erkannt/Generic
- `★` (U+2605) Eigene IDs

In Panel-Titeln und Buttons (Segoe UI) sind farbige Emojis OK: `🔍 📋 💾 ➕ 🔄 ⏳`.

### Kategorien-System
- **Kuratiert**: Kritisch, Fehler, Warnung, Information, Überwachung
- **Erkannt**: aus Phase-1-Scan (real aufgetreten)
- **Dokumentiert**: aus Phase-2-Scan (im Manifest deklariert)
- **Eigene**: vom Benutzer manuell hinzugefügt

### Bereichs-Filter („cbLog")
Nicht alle Optionen im Dropdown sind 1:1 ein Log-Name. Die Sonderfälle werden in
`Update-EventList` per `switch` auf `$filterLog` behandelt. Beim Erweitern um neue
Optionen: dort den passenden `$showLog`-Zweig hinzufügen.

## Bekannte Einschränkungen

- **Remote-Computer**: erfordert Firewall-Regel "Remote-Ereignisprotokoll-Verwaltung"
  und Admin-Rechte auf dem Ziel
- **Sehr lange Manifest-Beschreibungen**: enthalten Platzhalter wie `%1`, `%2` —
  das ist normal (werden zur Laufzeit ersetzt)
- **Performance**: Phase-2-Scan bei sehr vielen Providern (>200) kann 30+ Sekunden dauern
- **Win11 Sicherheits-Auditing**: Viele Security-IDs (z.B. 4688) erscheinen nur, wenn
  die entsprechende Audit-Policy aktiviert ist

## Nächste Schritte / Ideen

- [ ] Mehrfach-Computer-Auswahl (Liste statt einzelnes Textfeld)
- [ ] Speichern/Laden von Auswahl-Profilen (z.B. „Anmelde-Audit", „Defender-Vorfälle")
- [ ] Live-Modus: periodisches Refresh der Ergebnisliste
- [ ] Excel-Export zusätzlich zu CSV
- [ ] Zeitreihe-Diagramm pro Event-ID im Ausgabeformular
- [ ] XPath-Direktabfrage als Power-User-Modus (umgeht das 22-IDs-Limit elegant)

## Coding-Konventionen

- WinForms-Objekte konsistent über `New-Object` aufbauen, NICHT mit `[Forms.X]::new()`
- Click-Handler als ScriptBlocks an `Add_Click`
- Strings auf Deutsch (UI-Sprache), Variable & Funktionen Englisch
- Skript läuft sowohl unter Windows PowerShell 5.1 als auch PowerShell 7+

## Test-Hinweise

Manuelle Testfälle, die immer wieder gebrochen sind:
1. **Auswahl von >22 IDs aus einem Log** — muss in Batches aufgeteilt werden
2. **Filter wechseln während CLB Items gecheckt sind** — Checks gehen verloren (akzeptiert)
3. **Re-Scan auf nicht erreichbaren Remote-Computer** — Statuszeile muss Fehler zeigen, kein Crash
4. **Eigene ID hinzufügen mit nicht-numerischem Text** — Validierung greift
5. **Doppelklick auf Ergebniszeile** — zeigt vollständige Nachricht im MessageBox
