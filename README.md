# Outlook-Font-Setter
Dieses Skript setzt in Outlook die Schriftart auf Verdana und die Schriftgröße auf 9. Es bearbeitet die entsprechenden Registry-Einträge, um die Standard-Schriftart für das Verfassen, Antworten und Weiterleiten von E-Mails zu ändern.
Funktionsweise

    Überprüft, ob der Registry-Pfad für die Mail-Einstellungen existiert, und erstellt diesen bei Bedarf.
    Setzt die Schriftart auf Verdana und die Schriftgröße auf 9 für verschiedene Textformate in Outlook.
    Enthält grundlegende Fehlerbehandlung, um sicherzustellen, dass Probleme beim Ändern der Registry-Einträge gemeldet werden.

Voraussetzungen

    Windows Betriebssystem
    PowerShell 5.0 oder höher
    Administratorrechte, um Änderungen an der Registry vorzunehmen

Verwendung

    Lade das Skript herunter und speichere es lokal auf deinem Computer.
    Öffne eine PowerShell-Konsole mit Administratorrechten.
    Navigiere zu dem Verzeichnis, in dem du das Skript gespeichert hast.
    Führe das Skript mit folgendem Befehl aus:

    powershell

    .\Set-OutlookFont-Verdana9.ps1

Registry-Schlüssel und Werte

Das Skript aktualisiert folgende Registry-Schlüssel:

    Template
    MarkCommentsWith
    TextFontComplex
    TextFontSimple
    ComposeFontComplex
    ComposeFontSimple
    ReplyFontComplex
    ReplyFontSimple
