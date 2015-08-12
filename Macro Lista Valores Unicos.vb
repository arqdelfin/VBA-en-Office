Sub CreaListaUNICA()
    ' generamos una linea por cada 50 elementos en la tabla creada y un fichero txt por cada linea
    Set ListaOrigen = Application.InputBox(Prompt:="Rango de datos origen:", Title:="Seleccionar rango", Type:=8)
    
    ' opcion de creacion de linea con los valores UNICOS de la seleccion
    Set Dict1 = CreateObject("scripting.dictionary")
    Set WSF = Application.WorksheetFunction
    
    With Dict1
        .comparemode = 1
        For Each rng In ListaOrigen
             If Dict1.Exists(rng.Text) Then
                 Dict1.Item(rng.Value) = Dict1.Item(rng.Value) + 1
                 Else
                 Dict1.Add rng.Text, 1
             End If
        Next rng
    End With
    
    nRows = Dict1.Count ' numero de lineas en el rango seleccionado
    If Dict1.Count > 200 Then
        MsgBox "El numero de elementos excede del maximo (200).", vbCritical, "Demasiados Elementos"
        Else
        nLine = Int(nRows / 50) ' calculo en numero de lineas con 50 elementos
        ' borro las celdas de salida
        Sheets("Hoja1").Range("A1:Q5").Clear
        ' formatea las lineas de salida
        Range("A1:Q1").Select
        Selection.Merge
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        Selection.Copy
        Range("A2").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Selection.Copy
        Range("A3").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Selection.Copy
        Range("A4").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Selection.Copy
        
        ' bucle por cada linea que genere
        For m = 0 To nLine
            ult = ((m + 1) * 50) - 1
            If ult >= nRows Then ult = (nRows - 1)
            ' bucle para generar la linea
            For n = (m * 50) To ult
                Rea = Rea & Dict1.Keys()(n) & " or "
            Next
            ' prepara la linea y la añade a la Hoja1
            Rea = Left(Rea, Len(Rea) - 4) & vbCrLf ' elimina el ultimo Or de la linea
            Sheets("Hoja1").Cells(m + 1, 1) = Rea
            Rea = vbNullString ' vacia el campo Rea
        Next
        Dict1.RemoveAll
        m = n = 0
        'GeneraTXT
    End If
End Sub