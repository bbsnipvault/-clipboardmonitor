Ein kleines VBScript, das kontinuierlich die Windows-Zwischenablage überwacht.
Sobald eine neue Zeichenkette erkannt wird, filtert das Skript ausschließlich die enthaltenen Ziffern heraus.

Ab einer definierten Startsequenz – standardmäßig 400 – wird der restliche Zahlenstring extrahiert, zurück in die Zwischenablage kopiert und automatisch per STRG+V in das aktive Fenster eingefügt.

Dieses Tool eignet sich ideal zur schnellen und automatisierten Weiterverarbeitung numerischer Daten, zum Beispiel aus:

E-Mails,Barcodes,Scans


Überblick:

Permanente Überwachung der Zwischenablage und automatisches Ausgeben der Zahlenfolgen.Automatisches Einfügen in das aktive Fenster (per STRG+V Funktion)
Extraktion nur der Ziffern ohne Sonderzeichen oder Buchstaben
Startsequenz (startPattern) frei konfigurierbar (Im script wird 400 verwendet)
