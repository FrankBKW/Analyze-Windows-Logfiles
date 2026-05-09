# Windows Event Analyzer

Interaktives PowerShell-GUI-Tool zum Abfragen, Filtern und Auswerten von Windows-Ereignisprotokollen – lokal und auf Remote-Computern.

---

## Inhaltsverzeichnis

1. [Voraussetzungen](#voraussetzungen)
2. [Installation & Start](#installation--start)
3. [Übersicht der Oberfläche](#übersicht-der-oberfläche)
4. [Abfrage-Optionen](#abfrage-optionen)
5. [Credentials für Remote-Zugriff](#credentials-für-remote-zugriff)
6. [Mehrfach-Computer](#mehrfach-computer)
7. [Auswahl-Profile](#auswahl-profile)
8. [Live-Modus](#live-modus)
9. [Event-Auswahl](#event-auswahl)
10. [Eigene Event-IDs](#eigene-event-ids)
11. [XPath-Direktabfrage](#xpath-direktabfrage)
12. [Ergebnis-Fenster](#ergebnis-fenster)
13. [Export](#export)
14. [Zeitreihe-Diagramm](#zeitreihe-diagramm)
15. [Bekannte Einschränkungen](#bekannte-einschränkungen)

---

## Voraussetzungen

| Anforderung | Details |
|---|---|
| **Betriebssystem** | Windows 10 / Windows 11 / Windows Server 2016+ |
| **PowerShell** | 5.1 (vorinstalliert) oder PowerShell 7+ |
| **Rechte** | Administrator-Rechte für vollständigen Log-Zugriff (Security-Log, etc.) |
| **Remote-Zugriff** | WinRM muss auf den Ziel-Computern aktiv sein (`Enable-PSRemoting`) |
| **Excel-Export** | Microsoft Excel muss installiert sein (COM-Automatisierung) |

---

## Installation & Start

1. **Script herunterladen** oder in ein beliebiges Verzeichnis legen:
   ```
   WindowsEventAnalyzer.ps1
   ```

2. **Als Administrator starten** – Rechtsklick auf PowerShell → *Als Administrator ausführen*:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "WindowsEventAnalyzer.ps1"
   ```

   Oder direkt in einer Admin-PowerShell:
   ```powershell
   & "C:\Pfad\zum\WindowsEventAnalyzer.ps1"
   ```

3. Beim Start wird automatisch ein **Computer-Scan** durchgeführt: Alle aktiven Event-Logs des lokalen Computers werden eingelesen und der Katalog um erkannte Event-IDs ergänzt. Dieser Vorgang dauert je nach System 10–60 Sekunden und kann mit **Überspringen** abgebrochen werden.

---

## Übersicht der Oberfläche

```
┌─────────────────────────────────────────────────────────────────────┐
│  🔍  Windows Event Abfrage-Tool                                     │ ← Titelleiste
├─────────────────────────────────────────────────────────────────────┤
│  ⚙ Abfrage-Optionen   [Zeitraum] [Max] [Computer] [🔄 Scan]        │
│                        [Domain] [Benutzer] [Passwort]               │ ← Optionen
│                        [📋 Profil] [📂 Laden] [💾 Speichern]       │
│                        [⟳ Live-Modus] [Intervall]                  │
├─────────────────────────────────────────────────────────────────────┤
│  📋 Events auswählen   [Kategorie ▼] [Bereich ▼] [✔ Alle] [✖ Keine]│
│  ☐ ● ID 41   [System    ] [Kritisch  ]  Kernel-Power: Unerwarteter  │
│  ☑ ▲ ID 4625 [Security  ] [Warnung   ]  Fehlgeschlagene Anmeldung   │
│  ☐ ○ ID 7036 [System    ] [Information] Dienst gestartet...         │
│  ...                                                                 │ ← Auswahlliste
├─────────────────────────────────────────────────────────────────────┤
│  ➕ Eigene Event-ID hinzufügen                                      │
├─────────────────────────────────────────────────────────────────────┤
│  [Beenden]  [Status...]  [🧪 XPath-Abfrage]  [🔎 Abfragen]        │ ← Buttons
└─────────────────────────────────────────────────────────────────────┘
```

---

## Abfrage-Optionen

### Zeitraum
Wählt den Zeitraum, aus dem Ereignisse geladen werden:

| Auswahl | Beschreibung |
|---|---|
| Letzte 1 Stunde | Ereignisse der letzten 60 Minuten |
| Letzte 6 Stunden | Ereignisse der letzten 6 Stunden |
| Letzte 24 Stunden | *(Standard)* Ereignisse des letzten Tages |
| Letzte 7 Tage | Ereignisse der letzten Woche |
| Letzte 30 Tage | Ereignisse des letzten Monats |
| Alles | Kein Zeitfilter – alle vorhandenen Einträge |

> **Tipp:** Bei großen Logs oder vielen ausgewählten Events kann „Alles" sehr lange dauern. Mit **Max. Einträge** lässt sich die Menge begrenzen.

### Max. Einträge
Begrenzt die Anzahl der zurückgelieferten Ereignisse **pro Log-Abfrage**:
`25 / 50 / 100 / 250 / 500`

### Computer
Gibt an, welcher Computer abgefragt wird. Leer lassen = lokaler Computer.

---

## Credentials für Remote-Zugriff

Um Ereignisprotokolle eines anderen Computers auszulesen, müssen in der Regel Anmeldedaten angegeben werden.

| Feld | Beschreibung |
|---|---|
| **Domain** | Active-Directory-Domäne (z.B. `FIRMA`). Leer lassen für lokale Konten. |
| **Benutzer** | Benutzername (z.B. `Administrator`) |
| **Passwort** | Wird beim Aufruf als `SecureString` behandelt und nicht gespeichert. |

> Sind alle drei Felder leer, wird die **aktuelle Windows-Anmeldung** verwendet (Kerberos / Pass-Through-Auth).

---

## Mehrfach-Computer

Im Feld **Computer** können mehrere Hostnamen oder IP-Adressen kommagetrennt eingegeben werden:

```
PC-EMPFANG, PC-BUCHHALTUNG, SRV-FILESERVER01
```

- Die Abfrage wird für **jeden Computer separat** ausgeführt.
- Im Ergebnis-Fenster erscheint eine zusätzliche Spalte **Computer** sowie ein Filter-Dropdown zum Einschränken auf einen einzelnen Host.
- Die angegebenen Credentials (Domain/Benutzer/Passwort) gelten für alle Computer gleichermaßen.

> **Voraussetzung Remote:** WinRM muss auf den Zielcomputern aktiv sein:
> ```powershell
> Enable-PSRemoting -Force   # auf jedem Ziel-Computer einmalig ausführen
> ```

---

## Auswahl-Profile

Profile speichern die aktuelle Event-Auswahl (welche IDs angehakt sind), den Zeitraum und die Max-Einträge – damit häufig verwendete Abfrage-Konfigurationen schnell wiederhergestellt werden können.

### Profil speichern
1. Die gewünschten Events in der Liste **anhaken**.
2. Im Feld **📋 Profil** einen Namen eingeben (z.B. `Security-Audit`).
3. **💾 Speichern** klicken.

Das Profil wird als JSON-Datei gespeichert unter:
```
%APPDATA%\WindowsEventAnalyzer\profiles\Security-Audit.json
```

### Profil laden
1. Im Dropdown **📋 Profil** das gewünschte Profil auswählen.
2. **📂 Laden** klicken.

Die Haken in der Event-Liste werden automatisch gesetzt. Einträge, die im aktuellen Filter nicht sichtbar sind, werden beim Laden auf „Alle Bereiche" umgeschaltet.

---

## Live-Modus

Aktiviert ein automatisches, periodisches Aktualisieren der Ergebnisliste im Ausgabe-Fenster.

### Einrichtung
1. Checkbox **⟳ Live-Modus** aktivieren.
2. Im Dropdown das gewünschte **Intervall** wählen:
   - `30 Sek` · `1 Minute` *(Standard)* · `5 Minuten` · `10 Minuten`
3. **🔎 Abfragen** klicken – das Ausgabe-Fenster öffnet sich.

Im Ausgabe-Fenster erscheint unten ein Statustext mit dem Zeitstempel der letzten Aktualisierung:
```
⟳ Live: letzte Aktualisierung 14:37:52  –  42 Einträge
```

> Der Timer läuft nur so lange, wie das Ausgabe-Fenster geöffnet ist. Beim Schließen stoppt er automatisch.

---

## Event-Auswahl

### Kategorie-Filter
Zeigt nur Events einer bestimmten Schwere:
`Alle Kategorien / Kritisch / Fehler / Warnung / Information / Überwachung / Eigene / Erkannt / Dokumentiert`

### Bereichs-Filter
Schränkt die Anzeige auf eine Log-Gruppe ein:

| Auswahl | Beschreibung |
|---|---|
| Alle Bereiche | Kompletter Katalog |
| 🔍 Gefunden (alle Scan-Ergebnisse) | Alle beim Startup-Scan erkannten IDs |
| ★ Empfohlen (kuratiert) | Vordefinierter Katalog mit 120+ wichtigen IDs |
| ◈ Erkannt auf Computer | Nur IDs, die tatsächlich im Sample vorkamen |
| 📚 Dokumentiert (Manifest) | IDs aus Provider-Manifesten |
| System / Security / Application / … | Einzelne Log-Gruppen |

### Mehrfachauswahl
- **Einzeln** per Mausklick auf die Checkbox in der Liste.
- **✔ Alle** / **✖ Keine** – alle sichtbaren Einträge an-/abhaken.
- **Profil laden** – vordefinierte Auswahl wiederherstellen.

### Hover-Info
Ein Klick auf einen Listeneintrag zeigt unten die vollständige Beschreibung der Event-ID an.

### Scan-Aktualisierung (🔄 Scan)
Der **🔄 Scan**-Button fragt den im Computer-Feld eingetragenen Host erneut ab und ergänzt den Katalog um dort gefundene Event-IDs. Nützlich, wenn der Remote-Computer zuerst nicht erreichbar war.

---

## Eigene Event-IDs

Im Abschnitt **➕ Eigene Event-ID hinzufügen** können IDs manuell ergänzt werden, die nicht im Standard-Katalog enthalten sind.

| Feld | Beschreibung |
|---|---|
| **Event-ID** | Numerische ID (nur Ziffern) |
| **Log/Protokoll** | Dropdown mit Vorschlägen + frei editierbar |
| **Beschreibung** | Optionaler Freitext |

Nach dem Klick auf **➕ Hinzufügen** erscheint die ID sofort in der Liste und wird automatisch angehakt. Den genauen Log-Namen findet man in der Windows-Ereignisanzeige unter *Eigenschaften des Logs*.

---

## XPath-Direktabfrage

Für Experten: vollständig freier Filter über die XPath-Syntax der Windows-Ereignisanzeige.

**Öffnen:** Button **🧪 XPath-Abfrage** in der unteren Buttonleiste.

### Felder im Dialog

| Feld | Beschreibung |
|---|---|
| **Log-Name** | Name des Ereignisprotokolls (Dropdown + frei editierbar) |
| **Computer** | Ziel-Computer (vorbelegt aus dem Hauptformular) |
| **Max. Einträge** | Begrenzung der Ergebnismenge |
| **XPath-Filter** | Beliebiger XPath-Ausdruck |

### Beispiele

Alle kritischen und Fehler-Ereignisse der letzten 24 Stunden:
```xpath
*[System[(Level=1 or Level=2) and TimeCreated[timediff(@SystemTime) <= 86400000]]]
```

Alle fehlgeschlagenen Anmeldungen der letzten Stunde:
```xpath
*[System[EventID=4625 and TimeCreated[timediff(@SystemTime) <= 3600000]]]
```

Bestimmter Benutzer + Event-ID:
```xpath
*[System[EventID=4624] and EventData[Data[@Name='TargetUserName']='Max.Mustermann']]
```

> Die Ergebnisse werden in einem eigenen Ausgabe-Fenster mit Freitext-Filter und CSV-Export geöffnet.

---

## Ergebnis-Fenster

Nach dem Klick auf **🔎 Abfragen** öffnet sich das Ausgabe-Fenster.

### Spalten

| Spalte | Beschreibung |
|---|---|
| **Computer** | Quell-Computer des Ereignisses |
| **Zeit** | Zeitstempel (TT.MM.JJJJ HH:mm:ss) |
| **Typ** | Critical / Error / Warning / Information / Verbose |
| **Protokoll** | Kurzname des Event-Logs |
| **EventID** | Numerische Ereignis-ID |
| **Kategorie** | Aus dem Katalog (Kritisch, Warnung, Überwachung, …) |
| **Quelle** | Provider-Name |
| **Beschreibung** | Katalogbeschreibung der Event-ID |
| **Nachricht** | Erste 300 Zeichen der Ereignismeldung |

### Zeilenfarben

| Farbe | Bedeutung |
|---|---|
| 🔴 Rot | Critical |
| 🩷 Lachs | Error |
| 🟡 Gelb | Warning |
| Weiß | Information / Verbose |

### Filter in der Ergebnisliste

- **🔎 Freitext** – durchsucht EventID, Quelle, Nachricht und Beschreibung gleichzeitig.
- **Typ** – filtert auf einen Schweregrad.
- **Protokoll** – filtert auf einen einzelnen Log.
- **Computer** – filtert auf einen einzelnen Host (bei Mehrfach-Abfragen).

### Detailansicht
- **Einfachklick** auf eine Zeile → Kurzansicht der Nachricht in der Statuszeile unten.
- **Doppelklick** → Vollständige Ereignismeldung inkl. aller Felder in einem Popup.

---

## Export

### CSV-Export (💾 CSV)
Exportiert **alle** aktuell abgefragten Ergebnisse (nicht nur die gefilterte Ansicht) als semikolon-getrennte CSV-Datei (`UTF-8`).

### Excel-Export (📊 Excel)
Exportiert die Ergebnisse als formatierte `.xlsx`-Datei:
- Kopfzeile fett + farbig hinterlegt
- Zeilenfärbung nach Schweregrad (Rot / Lachs / Gelb)
- Automatische Spaltenbreite (Nachricht-Spalte max. 80 Zeichen)

> **Voraussetzung:** Microsoft Excel muss installiert sein. Das Tool nutzt COM-Automatisierung (`Excel.Application`).

---

## Zeitreihe-Diagramm

Zeigt die Event-Häufigkeit über die Zeit als Balkendiagramm.

**Öffnen:** Button **📈 Diagramm** im Ergebnis-Fenster.

- Bei einem Zeitraum **≤ 48 Stunden** → Granularität **pro Stunde**
- Bei einem Zeitraum **> 48 Stunden** → Granularität **pro Tag**

Das Diagramm berücksichtigt **alle** abgefragten Ergebnisse (vor dem Filter). Es gibt einen schnellen Überblick, wann Ereignisse gehäuft aufgetreten sind.

> **Voraussetzung:** `System.Windows.Forms.DataVisualization` (.NET Framework 3.5+, auf modernen Windows-Systemen standardmäßig vorhanden).

---

## Bekannte Einschränkungen

| Thema | Details |
|---|---|
| **22-ID-Limit** | `Get-WinEvent -FilterHashtable` unterstützt maximal 22 Event-IDs pro Aufruf. Das Tool teilt die Auswahl automatisch in Batches auf – dadurch ist `Max. Einträge` ein Limit *pro Batch*, nicht über alle IDs. |
| **Security-Log** | Erfordert **lokale Administratorrechte** oder Mitgliedschaft in der Gruppe *Event Log Readers*. |
| **Remote WinRM** | Muss auf Ziel-Computern aktiviert sein. Firewall-Ausnahmen für Port 5985 (HTTP) oder 5986 (HTTPS) notwendig. |
| **Excel nicht vorhanden** | Ohne installiertes Excel steht nur CSV-Export zur Verfügung. |
| **Diagramm-Assembly** | Fehlt `System.Windows.Forms.DataVisualization`, öffnet sich das Diagramm-Fenster mit einer Fehlermeldung. |
| **Startup-Scan** | Auf Systemen mit sehr vielen Logs (> 300) kann der initiale Scan mehrere Minuten dauern. Abbrechen mit **Überspringen** ist jederzeit möglich. |

---

## Profil-Dateien (Speicherort)

```
%APPDATA%\WindowsEventAnalyzer\profiles\<Profilname>.json
```

Beispielinhalt einer Profil-Datei:
```json
{
  "Name": "Security-Audit",
  "Zeitraum": 3,
  "MaxCount": 2,
  "Events": [
    { "ID": 4625, "Log": "Security" },
    { "ID": 4740, "Log": "Security" },
    { "ID": 4720, "Log": "Security" }
  ]
}
```

---

## Schnellstart-Beispiele

### Fehlgeschlagene Anmeldungen der letzten 24 Stunden
1. Bereich → `Security`
2. `ID 4625` (Fehlgeschlagene Anmeldung) anhaken
3. Zeitraum → `Letzte 24 Stunden`
4. **🔎 Abfragen**

### Systemabstürze der letzten 7 Tage
1. Bereich → `System`
2. `ID 41` (Kernel-Power), `ID 6008` (Unerwarteter Shutdown) anhaken
3. Zeitraum → `Letzte 7 Tage`, Max. Einträge → `250`
4. **🔎 Abfragen**

### Defender-Alarme auf mehreren Servern
1. Computer → `SRV01, SRV02, SRV03`
2. Domain / Benutzer / Passwort ausfüllen
3. Bereich → `Windows Defender`
4. Kategorie → `Kritisch` – nur kritische Meldungen anzeigen
5. Zeitraum → `Letzte 7 Tage`
6. **🔎 Abfragen** → im Ergebnis über **Computer**-Filter je Server filtern

---

*Erstellt mit PowerShell WinForms · Kein externes Modul erforderlich (außer Excel für xlsx-Export)*
