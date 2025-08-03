' =============================================================================
'  Clean400Watcher.vbs– Automatische Zwischenablage-Überwachung mit STRG+V
'
' Dieses Skript überwacht kontinuierlich den Inhalt der Zwischenablage.
' Sobald eine neue Zeichenkette erkannt wird, extrahiert es ausschließlich die
' darin enthaltenen Ziffern. Ab einer definierten Startsequenz – standardmäßig "400" –
' wird der restliche Zahlenstring herausgefiltert, zurück in die Zwischenablage kopiert
' und automatisch per STRG+V in das aktive Fenster eingefügt.
'
' Die Startsequenz kann bei Bedarf individuell angepasst werden (siehe Variable „startPattern“).
' Dieses Skript eignet sich ideal zur schnellen und automatisierten Weiterverarbeitung
' nummerischer Daten – beispielsweise aus E-Mails, Barcodes oder Scans.
'
' -----------------------------------------------------------------------------
' Hinweis zur Reaktionsgeschwindigkeit:
'
' Am Ende der Hauptschleife (Do...Loop) befindet sich die Zeile:
'     WScript.Sleep 1000
' Diese legt die Wartezeit zwischen den Prüfzyklen in Millisekunden fest.
'
' Standardwert: 1000 ms = 1 Sekunde
' Für eine schnellere Reaktion z. B. alle 0,5 Sekunden, ändern Sie die Zeile wie folgt:
'     WScript.Sleep 500
' Beachten Sie, dass bei sehr kurzen Wartezeiten die CPU stärker beansprucht wird.
' =============================================================================


' Initialisiert Shell- und Dateisystem-Objekte
Dim shell, fso, rawText, cleaned, result, i, ch, pos400, tempFile
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Variable zum Zwischenspeichern des letzten Zwischenablageinhalts
Dim lastContent : lastContent = ""

Do
    ' -----------------------------------------------
    ' 1. Zwischenablage lesen (über PowerShell-Umweg)
    ' clipboard.txt = temporäre Datei, in der der aktuelle
    ' Zwischenablage-Inhalt gespeichert wird
    ' -----------------------------------------------
    shell.Run "cmd /c powershell -command Get-Clipboard > clipboard.txt", 0, True
    WScript.Sleep 300 ' kurze Pause, damit Datei fertig geschrieben ist

    If fso.FileExists("clipboard.txt") Then
        If fso.GetFile("clipboard.txt").Size > 0 Then

            ' Zwischenablageinhalt lesen und Zeilenumbrüche entfernen
            rawText = fso.OpenTextFile("clipboard.txt", 1).ReadAll
            rawText = Replace(rawText, vbCrLf, "")

            ' Nur weitermachen, wenn der Inhalt sich verändert hat
            If rawText <> lastContent Then
                lastContent = rawText
                cleaned = ""

                ' -----------------------------------------------
                ' 2. Nur Ziffern aus dem Text extrahieren
                ' Alle Nicht-Zahlenzeichen werden entfernt
                ' -----------------------------------------------
                For i = 1 To Len(rawText)
                    ch = Mid(rawText, i, 1)
                    If Asc(ch) >= 48 And Asc(ch) <= 57 Then
                        cleaned = cleaned & ch
                    End If
                Next

                ' -----------------------------------------------
                ' 3. Nach der Startposition "400" suchen
                ' → Wenn du nach einer anderen Zahl starten willst,
                '    ändere einfach "400" in die gewünschte Ziffernfolge
                ' Beispiel: "123", "700", etc.
                ' -----------------------------------------------
                pos400 = InStr(cleaned, "400")

                If pos400 > 0 Then
                    ' Extrahiere alles ab der Fundstelle von "400"
                    result = Mid(cleaned, pos400)

                    ' -----------------------------------------------
                    ' 4. Neuer Text wird in Datei geschrieben,
                    ' dann zurück in Zwischenablage geladen
                    ' -----------------------------------------------
                    tempFile = "result.txt"
                    Set outFile = fso.CreateTextFile(tempFile, True)
                    outFile.Write result
                    outFile.Close

                    ' Schreibe Inhalt zurück in Zwischenablage
                    shell.Run "cmd /c type result.txt | clip", 0, True
                    WScript.Sleep 100 ' kleine Pause vor STRG+V

                    ' -----------------------------------------------
                    ' 5. Automatisches Einfügen per STRG+V (SendKeys)
                    ' Der bereinigte Text wird damit ins aktive Fenster eingefügt
                    ' -----------------------------------------------
                    shell.SendKeys "^v"
                End If
            End If
        End If
    End If

    ' Warte 1 Sekunde bis zum nächsten Durchlauf
    WScript.Sleep 1000
Loop
