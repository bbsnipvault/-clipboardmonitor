' =============================================================================
' StopAllScripts.vbs – Beendet alle aktiven VBScript-Prozesse im Hintergrund
'
' Dieses Skript durchsucht alle aktuell laufenden Prozesse auf dem System
' und beendet gezielt alle aktiven VBScript-Instanzen – sowohl mit dem
' Windows-Skriptinterpreter **wscript.exe** als auch mit der Konsolenvariante **cscript.exe**.
'
' Der Hauptzweck dieses Tools ist es, festhängende oder unbeabsichtigt
' weiterlaufende Skripte zuverlässig zu stoppen – beispielsweise wenn ein
' Überwachungsskript wie „Clean400Watcher.vbs“ nicht mehr gebraucht wird.
'
' -----------------------------------------------------------------------------
' Funktionsweise:
'
' 1. Über die WMI-Schnittstelle („Windows Management Instrumentation“) wird eine
'    Abfrage an die Prozessliste gesendet:
'       "Select * from Win32_Process Where Name='wscript.exe' or Name='cscript.exe'"
'
' 2. Für jeden gefundenen Prozess mit dem Namen „wscript.exe“ oder „cscript.exe“
'    wird die Methode `.Terminate` aufgerufen – der Prozess wird sofort beendet.
'
' -----------------------------------------------------------------------------
' Optionaler Schutz vor unbeabsichtigtem Abbruch:
'
' In der auskommentierten Zeile im Code kann man einen bestimmten Skriptnamen prüfen:
'
'     If InStr(LCase(p.CommandLine), "clean400watcher.vbs") > 0 Then
'         p.Terminate
'     End If
'
' Das bedeutet: Nur Skripte, deren Befehlszeile (CommandLine) den Namen
' „clean400watcher.vbs“ enthält, werden gezielt beendet. So kann man verhindern,
' dass alle Skripte gestoppt werden – und stattdessen nur bestimmte gezielt abschießen.
'
' Um diesen Filter zu aktivieren, entferne einfach die Kommentarzeichen („'“) vor der `If`- und `End If`-Zeile.
' =============================================================================
