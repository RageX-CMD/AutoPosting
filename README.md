# Event-Post Generator

Windows-Anwendung zum Erstellen und optionalen Veröffentlichen der wöchentlichen
Naruto-Online-Eventbeiträge aus einer Excel-Datei.

Python oder eine Installation zusätzlicher Programme ist nicht erforderlich.

## Voraussetzungen

- Windows 10 oder Windows 11
- Microsoft Edge (empfohlen) oder Google Chrome
- Internetzugang
- Eine passende Excel-Datei mit Wochenblättern im Format `YYYYMMDD`
- Für automatisches Posting: ein Forum-Account mit Berechtigung für den
jeweiligen Bereich

## Welche EXE soll verwendet werden?

### EventPostGenerator_Allgemeines.exe

Veröffentlicht automatisch unter:

- Forumsbereich: **Allgemeines**
- Themenkategorie: **Aktionen**

Diese Version ist für normale Forum-Accounts vorgesehen.

### EventPostGenerator_News.exe

Veröffentlicht automatisch unter:

- Forumsbereich: **News-Bereich**
- Themenkategorie: **Aktionen**

Diese Version benötigt besondere Forum-Berechtigungen. Verwende sie nur, wenn
dein Account dort neue Themen erstellen darf.

Beide Programme können unabhängig voneinander verwendet werden und führen
getrennte Posting-Statusdateien.

## Excel-Datei

Die Excel-Datei kann:

- im gleichen Ordner wie die EXE liegen oder
- über **Durchsuchen** ausgewählt werden.

Jede Woche muss als eigenes Tabellenblatt im Format `YYYYMMDD` vorhanden sein,
zum Beispiel:

```text
20260917
20260924
```

Der Generator verwendet standardmäßig die neueste Woche und liest den deutschen
Text aus Spalte C. Bei älteren Tabellen kann der deutsche Text auch in einer
anderen der ersten drei Spalten erkannt werden.

## DOCX-Dateien erstellen

1. EXE starten.
2. Excel-Datei auswählen.
3. Gewünschte Woche auswählen.
4. Zielordner kontrollieren.
5. **DOCX erstellen** anklicken.

Für jede Woche wird ein eigener Ordner mit deutschem Datum erstellt:

```text
output/
└── 17.09.2026/
    ├── 20260917_Serverwartung_Aktionen.docx
    ├── 20260917_forum-post_01.docx
    ├── 20260917_forum-post_02.docx
    ├── 20260917_forum-post_03.docx
    └── 20260917_Kopieransicht.html
```

Die einzelnen `forum-post`-Dateien sind bereits passend zum Zeichenlimit des
Forums aufgeteilt.

## Alle fehlenden Wochen erstellen

Mit **Alle fehlenden Wochen** werden alle Excel-Wochen verarbeitet, für die noch
kein fertiger Wochenordner mit DOCX-Dateien existiert.

Bereits vorhandene Wochen werden nicht neu erzeugt oder überschrieben.

## Manuelles Copy-and-Paste

1. Woche über **DOCX erstellen** generieren.
2. **Kopieransicht öffnen** anklicken.
3. Im Browser bei jedem Segment **Beitrag formatiert kopieren** auswählen.
4. Inhalt direkt in den Forum-Editor einfügen.

Die Kopieransicht ist an den alten Forum-Editor angepasst und erhält:

- Fettschrift
- Überschriften
- Leerzeilen
- Event-Abstände
- das anklickbare Discord-Banner



## Automatisch posten

1. Richtige EXE für den gewünschten Forumsbereich verwenden.
2. Excel-Datei und Woche auswählen.
3. **Ausgewählte Woche automatisch posten** anklicken.
4. Zielbereich, Kategorie, Titel und Segmentanzahl kontrollieren.
5. Die Sicherheitsabfrage bestätigen.
6. Edge öffnet den passenden Forumsbereich.



### Erster Login

Beim ersten Mal:

1. Im geöffneten Edge selbst auf **Anmelden** klicken.
2. Login vollständig manuell durchführen.
3. In der App anschließend auf **OK** klicken.

Die App öffnet das Loginfenster absichtlich nicht selbst, um Cloudflare nicht
unnötig auszulösen.

Nach erfolgreichem Login sichert die App die Forum-Sitzung verschlüsselt unter:

```text
%LOCALAPPDATA%\EventPostGenerator\forum_session.dat
```

Die Datei:

- enthält keine lesbaren Cookie-Werte,
- ist mit Windows-DPAPI geschützt und
- kann nur vom gleichen Windows-Benutzer entschlüsselt werden.

Bei späteren Starts versucht die App, diese Sitzung automatisch
wiederherzustellen. Falls die Forumsseite die Sitzung ablaufen lässt, ist erneut
ein manueller Login erforderlich.

## Dauerhaftes Edge-Profil

Die App verwendet ein separates Edge-Profil:

```text
%LOCALAPPDATA%\EventPostGenerator\EdgeProfile
```

Mit `Forum-Profil_oeffnen.bat` kann dieses Profil auch ohne die App geöffnet
werden.

Wichtig: Alle Fenster dieses Profils müssen geschlossen sein, bevor die
Posting-App gestartet wird. Das Profil kann nicht gleichzeitig manuell und über
die App geöffnet sein.

## Ablauf beim automatischen Posting

- Das erste Segment wird als neues Thema mit dem Titel aus der Excel erstellt.
- Weitere Segmente werden als Antworten angehängt.
- Zwischen zwei Beiträgen wartet die App 35 Sekunden.
- Die App verwendet den echten Forum-Editor und den originalen
**Beitrag posten**-Button.
- Platzhalter wie `[Event-Bild hier einfügen]` werden sichtbar mitgepostet und
können anschließend bearbeitet werden.
- Der Browser bleibt nach Abschluss oder Fehlern zur Kontrolle geöffnet.



## Unterbrochenes Posting fortsetzen

Nach jedem erfolgreichen Segment wird im Wochenordner eine Statusdatei
gespeichert:

```text
posting_state.json
```

Für die News-Version:

```text
posting_state_news.json
```

Wird ein Lauf unterbrochen, fragt die App beim nächsten Versuch, ob bei der
nächsten fehlenden Antwort fortgesetzt werden soll.

Die Statusdatei verhindert außerdem, dass eine vollständig gepostete Woche
versehentlich doppelt veröffentlicht wird.

## Cloudflare-Block

Falls eine Seite mit **Sorry, you have been blocked** erscheint:

1. Posting-App und alle zugehörigen Edge-Fenster schließen.
2. Nicht mehrfach direkt hintereinander neu versuchen.
3. Einige Minuten warten.
4. Danach die App erneut starten und die Wiederaufnahme bestätigen.

Die App verwendet keine direkten Forum-API-Aufrufe. Beiträge werden über die
originalen Schaltflächen der Forumsseite abgesendet. Trotzdem kann Cloudflare
automatisierte Browser-Sitzungen im Einzelfall blockieren.

## Häufige Probleme



### Edge-Profil wird bereits verwendet

Alle Edge-Fenster schließen, die über `Forum-Profil_oeffnen.bat` oder die App
gestartet wurden. Anschließend erneut versuchen.

### Login wird nicht wiederhergestellt

Einmal über den von der App geöffneten Edge manuell anmelden und anschließend in
der App auf **OK** klicken. Nur dann kann die App die Sitzung verschlüsselt
sichern.

### 30-Sekunden-Fehlermeldung

Die aktuelle Version wartet 35 Sekunden zwischen Beiträgen. Prüfe, ob wirklich
die neueste EXE verwendet wird.

### Woche gilt bereits als gepostet

Die App hat eine vollständige `posting_state.json` beziehungsweise
`posting_state_news.json` gefunden. Vor einer manuellen Löschung prüfen, ob das
Thema tatsächlich bereits im Forum vorhanden ist.

### Windows SmartScreen warnt beim Start

Die EXE ist nicht digital signiert. Bei vertrauenswürdiger Herkunft kann unter
**Weitere Informationen** die Option **Trotzdem ausführen** gewählt werden.

## Weitergabe an andere Personen

Für die Nutzung reichen normalerweise:

- die passende EXE und
- die Excel-Datei.

Die andere Person benötigt kein Python. Sie muss sich beim ersten Posting mit
ihrem eigenen Forum-Account anmelden.

`forum_session.dat`, Edge-Profildaten und `posting_state`-Dateien sollten nicht
an andere Personen weitergegeben werden.