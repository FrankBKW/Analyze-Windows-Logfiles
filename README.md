# Windows Event Analyzer

![Version](https://img.shields.io/badge/Version-1.2.16-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![License](https://img.shields.io/badge/Lizenz-Privat%20%2F%20IT--intern-orange)

Interaktives PowerShell-GUI-Tool zum Abfragen, Filtern und Auswerten von Windows-Ereignisprotokollen – lokal und auf Remote-Computern.

| Eigenschaft   | Wert                        |
|---------------|-----------------------------|
| **Version**   | 1.2.16                       |
| **Datum**     | 2026-05-13                  |
| **Autor**     | FrankBKW                    |
| **Plattform** | Windows 10/11, Server 2016+ |
| **Sprache**   | PowerShell 5.1 / 7+         |

---

## ⚠️ Disclaimer / Rechtlicher Hinweis

> **Bitte vollständig lesen, bevor das Tool eingesetzt wird.**

### Nutzungszweck und Zielgruppe

Dieses Tool ist ausschließlich für den **legitimen, autorisierten Einsatz** durch IT-Administratoren, System-Engineers und Sicherheitsverantwortliche bestimmt, die zur Analyse der betroffenen Systeme **ausdrücklich berechtigt** sind. Jede andere Verwendung ist unzulässig.

### Zugriff auf fremde Systeme

Das Auslesen von Ereignisprotokollen eines fremden Computers ist in den meisten Ländern **strafbar**, wenn es ohne ausdrückliche Genehmigung des Systemeigentümers oder zuständigen IT-Verantwortlichen erfolgt.

- **Nur autorisierte Systeme abfragen.** Die Eingabe von Hostnamen oder Credentials, für die keine Berechtigung vorliegt, kann den Tatbestand des unbefugten Zugriffs auf ein Computersystem erfüllen (z.B. § 202a StGB in Deutschland, Art. 143 StGB in der Schweiz, § 118a öStGB in Österreich sowie die EU-Richtlinie 2013/40/EU und den Computer Fraud and Abuse Act in den USA).
- **Remote-Credentials sicher behandeln.** Eingegebene Passwörter werden im Arbeitsspeicher als `SecureString` verwaltet und nicht dauerhaft gespeichert. Dennoch liegt die Verantwortung für den sicheren Umgang mit Zugangsdaten beim Anwender.
- **Netzwerk-Policies beachten.** Der Einsatz in Unternehmensumgebungen muss mit den internen Sicherheitsrichtlinien und – sofern zutreffend – dem Betriebsrat oder der Personalvertretung abgestimmt sein.

### Datenschutz und personenbezogene Daten

Windows-Ereignisprotokolle können **personenbezogene Daten** enthalten, darunter:
- Benutzernamen, Anmeldezeiten und Arbeitsplatznamen
- IP-Adressen und Computernamen
- Datei- und Ressourcenzugriffe
- Prozessstart- und Anwendungsdaten

Der Umgang mit diesen Daten unterliegt der **Datenschutz-Grundverordnung (DSGVO / GDPR)** sowie den einschlägigen nationalen Datenschutzgesetzen. Exportierte Dateien (CSV, Excel) sind entsprechend zu schützen, zu kennzeichnen und nach Ablauf der gesetzlichen Aufbewahrungsfrist zu löschen.

### Ausführung als Administrator

Das Tool erfordert zur vollständigen Funktion lokale Administratorrechte. Die erhöhten Rechte sind **ausschließlich für den vorgesehenen Analysezweck** zu verwenden. Das Script ist vor der Ausführung auf die eigene Umgebung hin zu prüfen. Der Betreiber trägt die Verantwortung für den Einsatz auf seinen Systemen.

### Haftungsausschluss

- Dieses Tool wird **ohne jede Gewährleistung** bereitgestellt – weder ausdrücklich noch stillschweigend.
- Der Autor übernimmt **keine Haftung** für Schäden, die durch die Nutzung, Fehlnutzung oder den Ausfall des Tools entstehen, einschließlich Datenverlust, Systemausfälle oder Sicherheitsvorfälle.
- Die **Richtigkeit und Vollständigkeit** der angezeigten Ereignisdaten hängt von den Konfigurationen des Zielsystems ab. Das Tool trifft keine Aussagen über die Integrität der Protokolldaten.
- Der Einsatz in **Produktivumgebungen** erfolgt auf eigene Verantwortung. Es wird empfohlen, das Tool zunächst in einer Testumgebung zu erproben.

### Sicherheitshinweise zur Ausführungsrichtlinie

Der Start mit `-ExecutionPolicy Bypass` deaktiviert die PowerShell-Skriptausführungsrichtlinie für diesen Prozess. Dies ist ein bekanntes Sicherheitsrisiko. Empfehlung für den Produktiveinsatz:
```powershell
# Skript einmalig signieren (Code Signing Certificate erforderlich):
Set-AuthenticodeSignature -FilePath "WindowsEventAnalyzer.ps1" -Certificate $cert

# Oder Execution Policy nur auf den eigenen Benutzer anpassen:
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## Inhaltsverzeichnis

1. [Voraussetzungen](#voraussetzungen)
2. [Installation & Start](#installation--start)
3. [Sicherheitsanalyse](#sicherheitsanalyse)
4. [Übersicht der Oberfläche](#übersicht-der-oberfläche)
5. [Abfrage-Optionen](#abfrage-optionen)
6. [Credentials für Remote-Zugriff](#credentials-für-remote-zugriff)
7. [Mehrfach-Computer](#mehrfach-computer)
8. [Auswahl-Profile](#auswahl-profile)
9. [Live-Modus](#live-modus)
10. [Event-Auswahl](#event-auswahl)
11. [Eigene Event-IDs](#eigene-event-ids)
12. [XPath-Direktabfrage](#xpath-direktabfrage)
13. [Ergebnis-Fenster](#ergebnis-fenster)
14. [Export](#export)
15. [Zeitreihe-Diagramm](#zeitreihe-diagramm)
16. [Bekannte Einschränkungen](#bekannte-einschränkungen)

---

## Voraussetzungen

| Anforderung | Details |
|---|---|
| **Betriebssystem** | Windows 10 / Windows 11 / Windows Server 2016+ |
| **PowerShell** | 5.1 (vorinstalliert) oder PowerShell 7+ |
| **Rechte** | Administrator-Rechte für vollständigen Log-Zugriff (Security-Log, etc.) |
| **Remote-Zugriff** | WinRM muss auf den Ziel-Computern aktiv sein (`Enable-PSRemoting`) |
| **Excel-Export** | Optional – wird beim Start automatisch erkannt. Ohne Excel nur CSV-Export. |

---

## Installation & Start

### Option A – EXE (empfohlen, kein PowerShell-Wissen nötig)

1. **`WindowsEventAnalyzer.exe`** aus dem Release herunterladen.
2. **Doppelklick** auf die EXE – Windows fragt automatisch nach Administrator-Rechten (UAC-Dialog).
3. Fertig. Keine Installation, keine PowerShell-Konfiguration erforderlich.

> **Hinweis:** Da die EXE mit ps2exe aus dem PowerShell-Script gebaut wurde, kann Windows Defender oder ein Virenscanner beim ersten Start warnen. Dies ist ein bekanntes Verhalten bei mit ps2exe kompilierten Scripten. Die Quelldatei (`WindowsEventAnalyzer.ps1`) liegt offen zur Prüfung im Repository.

### Option B – PowerShell-Script (für Anpassungen / Entwicklung)

1. **`WindowsEventAnalyzer.ps1`** herunterladen oder in ein beliebiges Verzeichnis legen.

2. **Starten** – das Script erkennt fehlende Admin-Rechte automatisch und startet sich selbst per UAC-Dialog mit erhöhten Rechten neu:
   ```powershell
   powershell -ExecutionPolicy Bypass -File "WindowsEventAnalyzer.ps1"
   ```

   Alternativ direkt per Doppelklick auf die PS1 (sofern `.ps1`-Dateien mit PowerShell verknüpft sind) – die Auto-Elevation übernimmt den Rest.

   > **Hinweis:** Wer die PS1 bewusst ohne Admin-Rechte ausführen möchte (z. B. für Read-Only-Logs), kann den Elevation-Block am Dateianfang auskommentieren. Dann entfällt aber der Zugriff auf Security-Log und Remote-Computer.

3. Beim Start erscheint eine **Abfrage**, ob der lokale Computer gescannt werden soll:
   - **Ja** → Schnell-Scan läuft automatisch: Die aktivsten Event-Logs werden in wenigen Sekunden eingelesen und der Katalog um erkannte Event-IDs ergänzt. Der Vorgang kann jederzeit mit **Überspringen** abgebrochen werden.
   - **Nein** → Das Programm öffnet sofort ohne Scan. Ein Scan kann jederzeit manuell über den **🔄 Scan**-Button gestartet werden.

---

## Sicherheitsanalyse

Das Script wurde einer vollständigen statischen Sicherheitsanalyse unterzogen.

**Ergebnis: ✅ SICHER — kein schädlicher Code gefunden.**

| Prüfpunkt | Befund | Bewertung |
|---|---|---|
| **Backdoors / Exfiltration** | Kein WebClient, kein Socket, kein SMTP, keine ausgehenden Verbindungen zu externen Hosts | ✅ Sicher |
| **Credential-Handling** | Passwort wird sofort in `SecureString` konvertiert, nie in Dateien gespeichert, nie geloggt | ✅ Sicher |
| **Code-Injection** | Kein `Invoke-Expression`, kein `iex`, kein dynamisch ausgeführter Code | ✅ Sicher |
| **Persistenz** | Keine Registry-Schreibzugriffe, keine geplanten Tasks, kein Autostart | ✅ Sicher |
| **Dateioperationen** | Ausschließlich `%APPDATA%\WindowsEventAnalyzer\profiles` + benutzergesteuerte Exports via SaveFileDialog | ✅ Sicher |
| **Netzwerkzugriffe** | Nur `Get-WinEvent` auf explizit eingegebene Remote-Computer + Diagnosetests (Ping, Port 135) | ✅ Sicher |
| **Berechtigungs-Eskalation** | Standard-UAC-Elevation für Admin-Zugriff auf Event-Logs, kein Token-Missbrauch | ✅ Sicher |
| **Obfuskierung** | Vollständig lesbarer Klartext (3042 Zeilen), kein Base64, keine komprimierten Payloads | ✅ Sicher |

Der Quellcode (`WindowsEventAnalyzer.ps1`) liegt offen im Repository und kann jederzeit eingesehen werden.

---

## Übersicht der Oberfläche

```
┌─────────────────────────────────────────────────────────────────────┐
│  🔍  Windows Event Abfrage-Tool                                     │ ← Titelleiste
├─────────────────────────────────────────────────────────────────────┤
│  ⚙ Abfrage-Optionen   [Zeitraum] [Max] [Computer] [🔄 Scan] [□Manifest]  │
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
Gibt an, welcher Computer abgefragt wird. Leer lassen = lokaler Computer. Mehrere Hosts kommagetrennt eingeben.

---

## Credentials für Remote-Zugriff

Um Ereignisprotokolle eines anderen Computers auszulesen, müssen in der Regel Anmeldedaten angegeben werden.

| Feld | Beschreibung |
|---|---|
| **Domain** | Active-Directory-Domäne (z.B. `FIRMA`). Leer lassen für lokale Konten. |
| **Benutzer** | Benutzername (z.B. `Administrator`) |
| **Passwort** | Wird als `SecureString` behandelt und **nicht** dauerhaft gespeichert. |

> Sind alle drei Felder leer, wird die **aktuelle Windows-Anmeldung** verwendet (Kerberos / Pass-Through-Auth).

---

## Mehrfach-Computer

Im Feld **Computer** können mehrere Hostnamen oder IP-Adressen kommagetrennt eingegeben werden:

```
PC-EMPFANG, PC-BUCHHALTUNG, SRV-FILESERVER01
```

- Die Abfrage wird für **jeden Computer separat** ausgeführt.
- Im Ergebnis-Fenster erscheint eine zusätzliche Spalte **Computer** sowie ein Filter-Dropdown zum Einschränken auf einen einzelnen Host.
- Die angegebenen Credentials gelten für alle Computer gleichermaßen.

> **Voraussetzung Remote:** `Get-WinEvent -ComputerName` nutzt **kein WinRM**, sondern RPC/DCOM (Port 135).
> `Enable-PSRemoting` hilft hier **nicht**. Stattdessen auf dem **Ziel-Computer** als Admin ausführen:
> ```powershell
> # Firewall-Regel "Remote Event Log Management" aktivieren:
> Enable-NetFirewallRule -Name "RemoteEventLogSvc-In-TCP","RemoteEventLogSvc-NP-In-TCP","RemoteEventLogSvc-RPCSS-In-TCP"
>
> # Dienste sicherstellen:
> Get-Service EventLog, RemoteRegistry | Start-Service
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
| 📚 Dokumentiert (Manifest) | IDs aus Provider-Manifesten (nach Manifest-Scan) |
| System / Security / Application / … | Einzelne Log-Gruppen |

### Mehrfachauswahl
- **Einzeln** per Mausklick auf die Checkbox in der Liste.
- **✔ Alle** / **✖ Keine** – alle sichtbaren Einträge an-/abhaken.
- **Profil laden** – vordefinierte Auswahl wiederherstellen.

### Hover-Info
Ein Klick auf einen Listeneintrag zeigt unten die vollständige Beschreibung der Event-ID an.

### Fortschrittsanzeige bei Abfragen

Während eine Abfrage läuft, erscheint am unteren Rand des Hauptfensters ein **Fortschrittsbalken** mit aktuellem Status:
```
Abfrage läuft... PC-EXAMPLE  ·  Security  (2 / 5)
```
- Der **Abfragen**-Button wird während der Abfrage gesperrt, um Doppelklicks zu verhindern.
- Der Balken zeigt den Fortschritt pro Computer × Log-Gruppe.
- Nach Abschluss verschwindet der Balken automatisch.

### Lokaler Selbsttest (🖥️ Lokaler Test Button)

Der **Lokaler Test**-Button (untere Leiste) prüft auf dem aktuellen Computer warum Logs nicht gelesen werden können:

| Schritt | Was wird geprüft |
|---|---|
| 1) Administrator-Rechte | Läuft das Tool mit erhöhten Rechten? |
| 2) PowerShell / .NET | Version und Architektur |
| 3) EventLog-Dienst | Läuft der Windows EventLog-Dienst? |
| 4) ListLog * | Können alle Logs aufgelistet werden? |
| 5) System-Log | Direkter Lesezugriff |
| 6) Security-Log | Lesezugriff (erfordert Admin) |
| 7) Antivirus | Ist AV aktiv? (kann ps2exe-EXE blockieren) |

> **Hinweis:** Das Script ist **nicht** an einen bestimmten Computer gebunden. Alle Pfade und Namen werden dynamisch über `$env:COMPUTERNAME` und `$env:APPDATA` ermittelt. Funktioniert die EXE nicht, `.ps1` direkt starten: `powershell -ExecutionPolicy Bypass -File WindowsEventAnalyzer.ps1`

### Remote-Test (🔌 Remote-Test Button)

Der **Remote-Test**-Button (untere Leiste) prüft Schritt für Schritt warum ein Remote-Zugriff scheitert:

| Schritt | Was wird geprüft |
|---|---|
| 1) Ping | Netzwerk-Erreichbarkeit des Ziel-Computers |
| 2) Port 135 | RPC/DCOM-Port (benötigt für `Get-WinEvent`, **nicht** WinRM) |
| 3) Get-WinEvent | Direkter Zugriffstest mit exakter Fehlermeldung |
| 4) RemoteRegistry | Dienststatus auf dem Ziel |
| 5) Workgroup-UAC | Hinweis auf `LocalAccountTokenFilterPolicy` |

**Häufige Ursachen und Fixes:**

| Fehler | Fix auf Ziel-PC (als Admin) |
|---|---|
| Port 135 blockiert | `Enable-NetFirewallRule -Name "RemoteEventLogSvc-In-TCP","RemoteEventLogSvc-NP-In-TCP","RemoteEventLogSvc-RPCSS-In-TCP"` |
| Zugriff verweigert (Workgroup) | `reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f` → Neustart |
| Zugriff verweigert (Rechte) | Benutzer zur Gruppe *Event Log Readers* hinzufügen |
| RemoteRegistry gestoppt | `Get-Service RemoteRegistry \| Start-Service` |

### Scan-Diagnose bei Problemen

Findet der Scan **keine Event-IDs**, erscheint automatisch ein **Diagnose-Dialog** mit:
- Den genauen Fehlermeldungen pro Log
- Konkreten Fix-Befehlen je nach Fehlerart (lokal vs. remote)

**Lokal – keine Logs gefunden:**
- Tool als Administrator starten (Rechtsklick → *Als Administrator ausführen* / EXE zeigt UAC-Dialog automatisch)
- Antivirensoftware kann den Zugriff auf Event-Logs blockieren

**Remote – keine Logs gefunden:**
> `Get-WinEvent -ComputerName` nutzt **RPC/DCOM (Port 135)**, nicht WinRM.  
> `Enable-PSRemoting` hilft hier **nicht**.  
> Auf dem **Ziel-Computer** als Admin ausführen:
> ```powershell
> Enable-NetFirewallRule -Name "RemoteEventLogSvc-In-TCP","RemoteEventLogSvc-NP-In-TCP","RemoteEventLogSvc-RPCSS-In-TCP"
> Get-Service EventLog, RemoteRegistry | Start-Service
> ```

### Scan-Aktualisierung (🔄 Scan)
Der **🔄 Scan**-Button scannt den Ziel-Computer erneut und ergänzt den Katalog.

| Option | Verhalten |
|---|---|
| Checkbox **„Manifest"** nicht aktiv *(Standard)* | Schnell-Scan: nur aktiv aufgetretene IDs (5–15 Sek.) |
| Checkbox **„Manifest"** aktiv | Vollständiger Scan inkl. Provider-Manifesten – liefert auch nie aufgetretene dokumentierte IDs (kann mehrere Minuten dauern) |

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
| **Log-Name** | Name des Ereignisprotokolls (Dropdown + frei editierbar) – wird automatisch aus dem Bereichs-Filter des Hauptfensters vorbelegt (z.B. „Security" → `Security`, „PowerShell" → `Microsoft-Windows-PowerShell/Operational`) |
| **Computer** | Ziel-Computer – vorbelegt aus dem Hauptformular |
| **Max. Einträge** | Begrenzung der Ergebnismenge – vorbelegt aus dem Hauptformular |
| **XPath-Filter** | Beliebiger XPath-Ausdruck – wird automatisch mit dem gewählten Zeitraum aus dem Hauptfenster vorbelegt |

> **Automatische Voreinstellung:** Beim Öffnen des XPath-Dialogs werden Log-Name, Computer, Max. Einträge und der Zeitraum-Ausdruck im XPath-Filter direkt aus den aktuellen Einstellungen des Hauptfensters übernommen. Alle Felder bleiben manuell anpassbar.

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

> **Automatische Erkennung:** Beim Start des Tools wird geprüft, ob Microsoft Excel installiert ist. Ist Excel nicht vorhanden, wird der Excel-Export-Button in allen Ausgabefenstern automatisch ausgeblendet – es erscheint kein Fehler. Der CSV-Export steht immer zur Verfügung.

> **Datenschutz:** Exportierte Dateien können personenbezogene Daten enthalten. Zugriff schützen und nach der Auswertung sicher löschen.

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
| **22-ID-Limit** | `Get-WinEvent -FilterHashtable` unterstützt maximal 22 Event-IDs pro Aufruf. Das Tool teilt die Auswahl automatisch in Batches auf – `Max. Einträge` ist daher ein Limit *pro Batch*, nicht über alle IDs. |
| **Security-Log** | Erfordert **lokale Administratorrechte** oder Mitgliedschaft in der Gruppe *Event Log Readers*. |
| **Remote-Zugriff** | `Get-WinEvent -ComputerName` nutzt RPC/DCOM (Port 135), **nicht** WinRM. Auf dem Ziel als Admin ausführen: `Enable-NetFirewallRule -Name "RemoteEventLogSvc-In-TCP","RemoteEventLogSvc-NP-In-TCP","RemoteEventLogSvc-RPCSS-In-TCP"` (Regelname ist sprachunabhängig). `Enable-PSRemoting` allein reicht nicht. |
| **Excel nicht vorhanden** | Excel-Export-Button wird automatisch ausgeblendet. CSV-Export steht immer zur Verfügung. |
| **Diagramm-Assembly** | Fehlt `System.Windows.Forms.DataVisualization`, öffnet sich das Diagramm-Fenster mit einer Fehlermeldung. |
| **Nicht vorhandene Logs** | Logs die auf dem Ziel-Computer nicht existieren (z.B. „Windows Defender" auf einem Server) werden automatisch übersprungen – kein Fehler. Auch verwaiste Log-Registrierungen, deren Log-Datei auf dem Datenträger fehlt (z.B. RemoteFX-Debug), brechen den Scan nicht mehr ab – sie werden lautlos übersprungen. |
| **Scan-Sample-Größe** | Der Schnell-Scan liest nur 15 Events pro Log. IDs, die in keinem der letzten 15 Einträge eines Logs vorkommen, werden erst beim Manifest-Scan erkannt. |
| **Scan ohne WinRM (lokal)** | Der Scan erkennt automatisch ob der Ziel-Computer lokal ist und verzichtet dann auf `-ComputerName` (WinRM). Dadurch funktioniert der Scan auch ohne aktiviertes WinRM auf dem lokalen Rechner. |
| **Nachrichtentexte im Scan** | Der Startup-Scan zeigt keine Klartextnachrichten in der Erkannt-Liste (nur Provider-Name), da der Message-Lookup aus Performancegründen deaktiviert ist. Im Abfrage-Ergebnis erscheinen die vollständigen Nachrichten wie gewohnt. |
| **Startup-Scan optional** | Beim Start wird per Dialog gefragt, ob der lokale Computer gescannt werden soll. Bei Auswahl „Nein" enthält die Ereignisliste nur den Standard-Katalog – kein Eintrag unter „◈ Erkannt auf Computer". Scan jederzeit manuell nachholbar. |
| **Fenster im Vordergrund** | Alle Dialoge und Fenster (Startabfrage, Scan-Fortschritt, Ergebnisse, Diagnose, XPath, Diagramm) erscheinen immer im Vordergrund. Das Hauptfenster übergibt den Fokus nach dem Laden wieder an den Desktop. |

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
4. Kategorie → `Kritisch`
5. Zeitraum → `Letzte 7 Tage`
6. **🔎 Abfragen** → im Ergebnis über **Computer**-Filter je Server filtern

---

## Versionshistorie

| Version | Datum      | Highlights                                                                 |
|---------|------------|----------------------------------------------------------------------------|
| 1.2.16  | 2026-05-13 | Fix: Manifest-Checkbox-Label überragt nicht mehr den Panel-Rand                 |
| 1.2.15  | 2026-05-13 | Fix: op_Multiply endgültig behoben – alle * durch Addition/[int]-Typen ersetzt  |
| 1.2.14  | 2026-05-13 | Fix: Scanconfig erscheint beim Start vor der Scan/Überspringen-Auswahl          |
| 1.2.13  | 2026-05-13 | Fix: op_Multiply-Fehler behoben; Scanconfig automatisch vor erstem Abfragen     |
| 1.2.12  | 2026-05-13 | Filter-Zeile überschneidungsfrei (Resize-Handler); Beschreibung-Spalte entfernt; horizontale Scrollbar |
| 1.2.10  | 2026-05-13 | Scan-Einstellungen öffnen automatisch vor dem Scan                         |
| 1.2.9   | 2026-05-13 | Scan-Konfiguration: ⚙-Button mit Checkboxen pro Log-Gruppe + Scantiefe     |
| 1.2.8   | 2026-05-13 | Scan ohne MaxEvents-Limit (vollständige ID-Erkennung)                      |
| 1.2.7   | 2026-05-13 | Benutzer-Spalte in allen Ergebnissen (SID → Name aufgelöst)                |
| 1.2.5   | 2026-05-13 | Versionsnummer in Fenstern & EXE-Eigenschaften                             |
| 1.2.4   | 2026-05-13 | Fix: Beschreibung ≠ Nachricht (falscher Split auf `$short`); Fenster nicht mehr dauerhaft TopMost |
| 1.2.3   | 2026-05-13 | Fix: Beschreibung immer aus echter Event-Nachricht, nicht aus statischem Katalog |
| 1.2.2   | 2026-05-12 | Fix: Nachricht nicht mehr auf 300 Zeichen abgeschnitten                    |
| 1.2.1   | 2026-05-12 | Fix: Beschreibung im Live-Modus zeigte `[Provider:...]`-Metadaten          |
| 1.2.0   | 2026-05-12 | Major: Auto-Elevation PS1, alle Popups TopMost, Security-Review in README  |
| 1.1.x   | 2026-05-09 | Remote-Zugriff, XPath-Abfrage, Profil-System, Live-Modus, Diagramm        |
| 1.0.0   | 2026-05-09 | Erstveröffentlichung                                                       |

---

*Erstellt mit PowerShell WinForms · Kein externes Modul erforderlich (außer Microsoft Excel für xlsx-Export)*
