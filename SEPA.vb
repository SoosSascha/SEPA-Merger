Sub SEPACom()

Dim NextLine As Integer
Dim KdRow As Integer
Dim LastRow As Integer
Dim KdNummer As Integer
Dim KdName As String
Dim Betrag As Double
Dim BIC As String
Dim IBAN As String
Dim VZ As String
Dim Mref As String
Dim DatumMU As Date

GetWorkbook ("C:\Users\Buero\Desktop\SEPA.xlsx")

Set wb = ActiveWorkbook

Set wsl = Sheets("SEPAListe")
Set wsd = Sheets("SEPADaten")
Set wsz = Sheets("Ziel")

NextLine = wsl.Cells(wsl.Rows.Count, "E").End(xlUp).Row + 1

KdNummer = wsl.Cells(NextLine, 2)
VZ = wsl.Cells(NextLine, 3)
Betrag = wsl.Cells(NextLine, 4)

KdRow = wsd.Cells.Find(what:=KdNummer, LookAt:=xlWhole).Row

KdName = wsd.Cells(KdRow, 2)
BIC = wsd.Cells(KdRow, 4)
IBAN = wsd.Cells(KdRow, 5)
Mref = wsd.Cells(KdRow, 8)
DatumMU = wsd.Cells(KdRow, 9)

LastRow = wsz.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row + 1

wsz.Cells(LastRow, 1) = KdName
wsz.Cells(LastRow, 2) = Betrag
wsz.Cells(LastRow, 3) = BIC
wsz.Cells(LastRow, 4) = IBAN
wsz.Cells(LastRow, 5) = VZ
wsz.Cells(LastRow, 7) = Mref
wsz.Cells(LastRow, 8) = DatumMU

wsl.Cells(NextLine, 5) = "Angewiesen"

End Sub

Sub RowLoop()

GetWorkbook ("C:\Users\Buero\Desktop\SEPA.xlsx")

Set wb = ActiveWorkbook

Set wsl = Sheets("SEPAListe")
Set wsd = Sheets("SEPADaten")
Set wsz = Sheets("Ziel")

If wsl.Cells(wsl.Cells(wsl.Rows.Count, "E").End(xlUp).Row + 1, 1) = "" Then
Debug.Print "Keine neuen SEPA-Einzüge"
Else
Do Until wsl.Cells(wsl.Cells(wsl.Rows.Count, "E").End(xlUp).Row + 1, 1) = ""
    Call SEPACom
Loop
End If

Debug.Print "DONE"

End Sub

Public Function GetWorkbook(ByVal sFullName As String) As Workbook

    Dim sFile As String
    Dim wbReturn As Workbook

    sFile = Dir(sFullName)

    On Error Resume Next
        Set wbReturn = Workbooks(sFile)

        If wbReturn Is Nothing Then
            Set wbReturn = Workbooks.Open(sFullName)
        End If
    On Error GoTo 0

    Set GetWorkbook = wbReturn

End Function