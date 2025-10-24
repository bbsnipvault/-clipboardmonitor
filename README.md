1. CleanWatcher_400_or_60.vbs

Dieses Skript überwacht die Zwischenablage (Copy/Paste) auf deinem PC und prüft, ob neue Inhalte hineinkopiert wurden.

Zuerst wird nach der Zahl "400" gesucht.
→ Wird diese gefunden, werden die Ziffern ab dieser Stelle verarbeitet und auf maximal 12 Stellen begrenzt.

Falls keine "400" gefunden wird, wird nach "60" gesucht.
→ Wird dieser Wert gefunden, werden die Ziffern ab dieser Stelle genommen und auf maximal 13 Stellen begrenzt.

Ergebnisverarbeitung:

In die Datei result.txt geschrieben

In die Datei speicher.csv eingetragen (mit Datum und Uhrzeit)

In die Zwischenablage kopiert

Automatisch mit STRG+V nach einem definierten Zeitintervall ausgegeben

Wichtig:

Jede Zahl wird nur einmal verarbeitet (kein doppeltes Einfügen)

Die Überprüfung läuft in einer Endlosschleife im Hintergrund

Praktischer Nutzen:

Bereinigte Zahlenwerte werden direkt in einer CSV-Datei gespeichert, was den Zugriff auf Kassenzeichen erleichtert, insbesondere wenn Fachanwendungen keine eigene History-Funktion bieten.

2. RemoveSpacesAndLog.vbs

Dieses Skript entfernt automatisch alle Leerzeichen aus der Windows-Zwischenablage (Input Sanitization / Character Filtering) und aktualisiert die Zwischenablage.

Bereinigter Text wird zusätzlich mit Zeitstempel in history.txt gespeichert.

Optional kann der bereinigte Inhalt direkt per STRG+V eingefügt werden.

Praktischer Nutzen:

Nützlich, wenn Fachanwendungen diese Funktion nicht oder nur teilweise integriert haben.

Alle Skripte wurden mithilfe von KI erstellt und geprüft.
