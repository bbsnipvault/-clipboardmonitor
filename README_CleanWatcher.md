Clean400Watcher & StopAllScripts
================================

Dieses Repository enthält zwei einfache, nützliche Windows-VBS-Skripte zur Automatisierung mit der Zwischenablage und zur Kontrolle laufender VBScript-Prozesse.

Inhalt
------

Clean400Watcher.vbs
- Überwacht die Zwischenablage
- Entfernt alle Zeichen außer Zahlen
- Sucht ab einer definierten Startsequenz (z. B. "400")
- Kopiert den Rest in die Zwischenablage und fügt ihn per STRG+V automatisch ein

StopAllScripts.vbs
- Beendet alle laufenden VBScript-Prozesse (wscript.exe oder cscript.exe)
- Nützlich, um laufende oder hängen gebliebene Skripte zu beenden

Clean400Watcher.vbs
--------------------

Funktion:
Das Skript überprüft in regelmäßigen Abständen den Inhalt der Windows-Zwischenablage.
Wenn es einen neuen Inhalt erkennt, entfernt es alle Sonderzeichen und Buchstaben.
Es sucht dann nach einer festgelegten Startsequenz (z. B. "400"). Wird diese gefunden,
kopiert das Skript den Rest des Zahlenstrings in die Zwischenablage und fügt ihn automatisch
per STRG+V in das aktive Fenster ein.

Technische Details:
- Nutzt PowerShell, um den Zwischenablage-Inhalt in die Datei clipboard.txt zu schreiben
- Entfernt alle Zeichen außer Zahlen
- Startet die Verarbeitung ab der ersten Stelle, an der die definierte Startzeichenfolge vorkommt
- Kopiert die bereinigten Daten in result.txt und setzt den Inhalt zurück in die Zwischenablage
- Führt dann automatisch STRG+V aus (SendKeys "^v")

Geschwindigkeit einstellen:
Die Reaktionszeit wird über die Zeile

    WScript.Sleep 1000

gesteuert. Die Zahl steht für Millisekunden (1000 = 1 Sekunde).
Für eine schnellere Reaktion kann der Wert z. B. auf 500 reduziert werden:

    WScript.Sleep 500

Achtung: Kürzere Abstände können die CPU-Last erhöhen.

Startwert ändern:
Die Startsequenz "400" kann bei Bedarf im Code angepasst werden.
Die relevante Zeile lautet:

    pos400 = InStr(cleaned, "400")

Um z. B. ab "123" zu suchen:

    pos400 = InStr(cleaned, "123")

StopAllScripts.vbs
--------------------

Funktion:
Beendet alle laufenden VBScript-Prozesse (wscript.exe und cscript.exe).
Nützlich, wenn Skripte im Hintergrund weiterlaufen oder sich aufgehängt haben.

Optional: Nur gezielt Skripte stoppen
Die folgende auskommentierte Zeile kann aktiviert werden, um nur bestimmte Skripte zu beenden
(z. B. clean400watcher.vbs):

    If InStr(LCase(p.CommandLine), "clean400watcher.vbs") > 0 Then
        p.Terminate
    End If

Verwendung
----------

Clean400Watcher starten:
Doppelklick auf die Datei Clean400Watcher.vbs. Das Skript läuft dann im Hintergrund.

Alle Skripte stoppen:
Doppelklick auf StopAllScripts.vbs. Alle aktiven VBS-Prozesse werden beendet.

Dateien im Repository
----------------------

- Clean400Watcher.vbs – überwacht die Zwischenablage und reagiert automatisch
- StopAllScripts.vbs – beendet alle laufenden VBS-Prozesse
- README.md – diese Dokumentation

Voraussetzungen
---------------

- Windows 10 oder neuer
- Schreibrechte im Skriptordner (zum Speichern von clipboard.txt und result.txt)
- PowerShell muss aktiviert sein (Standard ab Windows 7)
- Windows Script Host muss aktiviert sein (Standard in Windows)

Lizenz
------

Veröffentlicht unter der MIT-Lizenz.
Freie Nutzung, Anpassung und Weitergabe sind erlaubt.
Verwendung auf eigene Verantwortung.

Hinweis
-------

KI ist ein Werkzeug – Verständnis und Verantwortung liegen beim Nutzer.