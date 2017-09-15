'Script per la colorazione di un'area come gradiente verticale per singola colonna
'Author Daniele Brambilla
'
'UTILIZZO:
'Impostare il colore iniziale nella prima cella della prima colonna dell'area a cui applicare il gradiente
'Impostare il colore finale nell'ultima cella della prima colonna dell'area a cui applicare il colore
'Selezionare l'area su cui applicare il gradiente
'Premere Alt + F11
'Dal menù 'Inserisci' selezionare la voce 'Modulo'
'Incollare il contenuto di questo file nella finestra che si apre
'Premere F5
'
'IMPORTANTE:
'È possibile applicare il gradiente solamente ad un'area alla volta, tuttavia l'area può essere arbitrariamente ampia
Sub applicaGradienteVerticalePerColonna()
    Dim selectedAreaText As String
    Dim redMask As Long
    Dim greenMask As Long
    Dim blueMask As Long
    Dim firstRed As Long
    Dim lastRed As Long
    Dim useRed As Long
    Dim firstGreen As Long
    Dim lastGreen As Long
    Dim useGreen As Long
    Dim firstBlue As Long
    Dim lastBlue As Long
    Dim useBlue As Long
    Dim factor As Double
    Dim firstColor As Long
    Dim lastColor As Long
    Dim useColor As Long
    Dim row As Long
    Dim column As Long
    Dim targetRange As Range
    Dim targetRowCount As Long
    Dim targetColumnCount As Long
    
    On Error Resume Next
    If ActiveWindow.RangeSelection.Count > 1 Then
      selectedAreaText = ActiveWindow.RangeSelection.AddressLocal
    Else
      selectedAreaText = ActiveSheet.UsedRange.AddressLocal
    End If
LInput:
    Set targetRange = Application.InputBox("Selezionare lo spazio delle celle su cui operare:", "Gradiente personalizzato", selectedAreaText, , , , , 8)
    If targetRange Is Nothing Then Exit Sub
    If targetRange.Areas.Count > 1 Then
        MsgBox "Selezioni multiple non supportate", vbInformation, "Gradiente personalizzato"
        GoTo LInput
    End If
    If targetRange Is Nothing Then Exit Sub
    If targetRange.Areas.Count > 1 Then
        MsgBox "Selezioni multiple non supportate", vbInformation, "Gradiente personalizzato"
        Exit Sub
    End If
    On Error Resume Next
    redMask = &HFF
    greenMask = &HFFFFFF And &HFF00FF00
    blueMask = &HFF0000
    Application.ScreenUpdating = False
    targetRowCount = targetRange.Rows.Count
    targetColumnCount = targetRange.Columns.Count
    For column = 1 To targetColumnCount
        firstColor = targetRange.Cells(1, column).Interior.Color
        firstRed = redMask And firstColor
        firstGreen = greenMask And firstColor
        firstBlue = blueMask And firstColor
        lastColor = targetRange.Cells(targetRowCount, column).Interior.Color
        lastRed = redMask And lastColor
        lastGreen = greenMask And lastColor
        lastBlue = blueMask And lastColor
        For row = 1 To targetRowCount
            factor = (row - 1) / (targetRowCount - 1)
            useRed = ((firstRed * (1 - factor)) + (lastRed * factor)) And redMask
            useGreen = ((firstGreen * (1 - factor)) + (lastGreen * factor)) And greenMask
            useBlue = ((firstBlue * (1 - factor)) + (lastBlue * factor)) And blueMask
            useColor = useRed Or (useGreen Or useBlue)
            targetRange.Cells(row, column).Interior.Color = useColor
        Next
    Next
End Sub


