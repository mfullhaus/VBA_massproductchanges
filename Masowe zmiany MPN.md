# VBA_massproductchanges
Masowe zmiany w produktach w katalogu - zmiany Kodu producenta
dodaj nowa zakładkę 'Sheet1' i nazwij kolumny 'ID' oraz "Kod"

	Sub DopasujKodProducenta()
    Dim wsLookup As Worksheet
    Dim ws As Worksheet
    Dim dict As Collection
    Dim lastRowLookup As Long
    Dim lastRow As Long
    Dim i As Long, r As Long
    Dim idValue As String
    Dim kodValue As Variant
    Dim prodCol As Long
    Dim cell As Range
    Dim idColLookup As Long, kodColLookup As Long
    Dim matchedKod As Variant

    ' Optymalizacja wydajności
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Ustawienie arkusza źródłowego "Sheet1"
    On Error Resume Next
    Set wsLookup = ThisWorkbook.Worksheets("Sheet1")
    On Error GoTo ErrorHandler

    If wsLookup Is Nothing Then
        MsgBox "Nie znaleziono arkusza 'Sheet1'!", vbCritical, "Błąd"
        GoTo CleanUp
    End If

    ' Natywna kolekcja VBA wolna od błędów ActiveX
    Set dict = New Collection

    ' Domyślne kolumny dla ID i Kodu w Sheet1
    idColLookup = 1
    kodColLookup = 2
    
    ' Dynamiczne wyszukiwanie kolumn ID oraz Kod w wierszach 1-2 arkusza Sheet1
    For i = 1 To 2
        Set cell = wsLookup.Rows(i).Find(What:="ID", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not cell Is Nothing Then idColLookup = cell.Column
        
        Set cell = wsLookup.Rows(i).Find(What:="Kod", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not cell Is Nothing Then kodColLookup = cell.Column
    Next i

    ' Wczytanie danych z Sheet1 do kolekcji
    lastRowLookup = wsLookup.Cells(wsLookup.Rows.Count, idColLookup).End(xlUp).Row
    For i = 2 To lastRowLookup
        idValue = Trim(CStr(wsLookup.Cells(i, idColLookup).Value))
        kodValue = wsLookup.Cells(i, kodColLookup).Value
        
        If idValue <> "" Then
            ' Ignorowanie ewentualnych duplikatów ID w Sheet1
            On Error Resume Next
            dict.Add Item:=kodValue, Key:=idValue
            On Error GoTo ErrorHandler
        End If
    Next i

    ' Przetwarzanie pozostałych arkuszy
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsLookup.Name Then
            ' Szukanie nagłówka "kod producenta" w wierszu 2
            Set cell = ws.Rows(2).Find(What:="kod producenta", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            
            If Not cell Is Nothing Then
                prodCol = cell.Column
                lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                
                ' Dopasowanie od wiersza 3 do końca danych
                For r = 3 To lastRow
                    idValue = Trim(CStr(ws.Cells(r, 1).Value))
                    If idValue <> "" Then
                        ' Pobranie wartości z kolekcji na podstawie ID
                        Err.Clear
                        On Error Resume Next
                        matchedKod = dict.Item(idValue)
                        If Err.Number = 0 Then
                            ws.Cells(r, prodCol).Value = matchedKod
                        End If
                        On Error GoTo ErrorHandler
                    End If
                Next r
            End If
        End If
    Next ws

    MsgBox "Dopasowanie kodów zostało zakończone pomyślnie!", vbInformation, "Sukces"

	CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

	ErrorHandler:
    MsgBox "Wystąpił błąd: " & Err.Description, vbCritical, "Błąd"
    GoTo CleanUp
	End Sub
