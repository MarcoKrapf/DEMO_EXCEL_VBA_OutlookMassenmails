Attribute VB_Name = "Modul1"
Option Explicit

Sub OutlookIndividualMassMail()
'Der Code durchl�uft die Spalte B des ersten Tabellenblatts
'mit den E-Mail-Adressen so lange, bis er keinen Eintrag
'mehr findet und schreibt die Namen aus Spalte A sowie die
'Mailadressen aus Spalte B in ein Array.
'Dann wird das Array durchlaufen und f�r jeden Eintrag eine
'E-Mail in Outlook generiert mit der E-Mail-Adresse im "An"-Feld,
'der pers�nlichen Anrede im Betreff und einem Text im Mail-Body.
'Die E-Mails werden entweder nur angelegt und ge�ffnet oder
'gleich versendet.

    Dim objOutlook As Object 'Variable f�r die Outlook-Applikation
    Dim objMail As Object 'Varable f�r die E-Mail
    Dim wks As Worksheet 'Variable f�r das Tabellenblatt
    Dim strRecipients() As String 'Array f�r die Aufnahme der Namen und Mailadressen
    Dim intLastRow As Integer 'Variable f�r den Wert der letzten Zeile
    Dim i As Integer, j As Integer 'Z�hlvariablen f�r Schleifen
    
    'Zuweisen des ersten Tabellenblatts dieser Excel-Mappe
    Set wks = ThisWorkbook.Sheets(1)
    
    'Letzte gef�llte Zelle in Spalte B ermitteln
    intLastRow = wks.Cells(Rows.Count, 2).End(xlUp).Row
    
    'Array neu dimensionieren: Zwei Spalten und
    'so viele Zeilen wie Eintr�ge in Spalte B
    ReDim strRecipients(1 To intLastRow, 1 To 2)
    
    'Tabellenblatt auslesen und Werte in das Array schreiben
    For i = 1 To UBound(strRecipients)
        For j = 1 To 2
            strRecipients(i, j) = wks.Cells(i, j).Value
        Next j
    Next i
    
    'Array durchlaufen und f�r jede Zeile eine E-Mail generieren
    For i = 1 To UBound(strRecipients)
    
        Set objOutlook = CreateObject("Outlook.Application")
        Set objMail = objOutlook.createitem(0) 'E-Mail erstellen
        
        With objMail
            .To = strRecipients(i, 2) '"An"
            .Subject = "Hallo " & strRecipients(i, 1) 'Betreff
            .body = "...E-Mail for " & strRecipients(i, 1) 'Mail-Body
            
            'Hier wird entschieden was gemacht werden soll
            .display 'Die Display-Methode �ffnet die E-Mail
                    'in Outlook, der Versand erfolgt anschlie�end manuell
'            .send 'Die Send-Methode sendet die E-Mail automatisch
        End With
        
    Next i
    
End Sub
