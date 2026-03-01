Attribute VB_Name = "kooste"
Option Explicit
Const rapo = "rapo_nimi"
Const kansio = "kansio"
Const jatke = ".xlsx"
Const piirit = "piirit"
Const laitteet = "laitteet"
Const viikkoja = 52
Const piireja = 5
Const laitteita = 3

Dim myynnit(viikkoja, piireja, laitteita) As Long
Dim piiri As Integer 'piirin nro
Dim laite As Integer 'laitteen nro
Dim viikko As Integer 'viikon nro
Dim tdo As String
Sub keraa_kaikki()

If Not tiedot_ok() Then
    MsgBox ("Tarkista tiedostot")
Else
    lue
    kirjoita
    raporttiin
End If

End Sub

Function tiedot_ok() As Boolean

If Dir(Range(kansio).Value, vbDirectory) = "" Then
    tiedot_ok = False
Else
    tiedot_ok = True
End If

End Function
Function lue()

For piiri = 1 To piireja
    Workbooks.Open (Range(kansio).Value & Range(piirit).Cells(piiri, 1).Value & jatke)
    For viikko = 1 To viikkoja
        For laite = 1 To laitteita
            myynnit(viikko, piiri, laite) = Cells(viikko + 1, laite + 1).Value
        Next
    Next
    ActiveWorkbook.Close
Next

End Function
Function kirjoita()

Dim r As Integer
r = 2
For viikko = 1 To viikkoja
    For piiri = 1 To piireja
        For laite = 1 To laitteita
            Cells(r, 1).Value = viikko
            Cells(r, 2).Value = Range(piirit).Cells(piiri, 1).Value
            Cells(r, 3).Value = Range(laitteet).Cells(laite, 1).Value
            Cells(r, 4).Value = myynnit(viikko, piiri, laite)
            r = r + 1
        Next
    Next
Next

End Function
Function raporttiin()

Dim pivotti As PivotTable
Worksheets(Range(rapo).Value).Activate

For Each pivotti In ActiveSheet.PivotTables
    pivotti.RefreshTable
Next pivotti

Range("A1").Select

End Function

