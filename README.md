This script sets the default font in Outlook to Verdana with a font size of 9. It modifies the corresponding registry entries to change the default font for composing, replying, and forwarding emails.
How It Works

    Checks if the registry path for mail settings exists and creates it if necessary.
    Sets the font to Verdana and the font size to 9 for various text formats in Outlook.
    Includes basic error handling to ensure any issues with modifying the registry entries are reported.

Prerequisites

    Windows operating system
    PowerShell 5.0 or higher
    Administrator privileges to make changes to the registry

Usage

    Download the script and save it locally on your computer.
    Open a PowerShell console with administrator privileges.
    Navigate to the directory where you saved the script.
    Run the script with the following command:

    powershell:

    .\Set-OutlookFont-Verdana9.ps1

Registry Keys and Values

The script updates the following registry keys:

    Template
    MarkCommentsWith
    TextFontComplex
    TextFontSimple
    ComposeFontComplex
    ComposeFontSimple
    ReplyFontComplex
    ReplyFontSimple

Customizing the Font

To customize the font used by the script, you need to manually set the desired font in Outlook first. Then, extract the registry entries (a .reg file) and save them to your desktop. Convert these entries using a RegtoXML script (available online or on GitHub). You can then replace the values in this script with the converted data.

This way, you do not need to use the default Aptos font.



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
