Sub Prueba()
Dim Ambito As Range

'Desactivamos servicios
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False

'Definimos el numero de filas de cada tabla (siempre menos de 65536)
ultArk = Sheets("Arkamax").Range("I65536").End(xlUp).Row
ultLib = Sheets("PacoBlanco").Range("A65536").End(xlUp).Row

'Cambiamos la barra de estado para que muestre nuestra informacion
oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True

'Asiganmos valores a Variables
Set Ambito = Sheets("PacoBlanco").Range("I2:I" & ultLib)
Datos = Ambito.Count
Id_Color = 36

Inicio = Now() 'Establecemos el incio del proceso

'Empieza el proceso de Busqueda y comparacion de registros entre las dos tablas.
    For X = 2 To ultArk
        Libro = Sheets("Arkamax").Cells(X, 9) 'Define el Titulo del Libro que vamos a buscar
        Application.StatusBar = "Se esta comprobando si los libros existen...... Espere" 'Quedan " & (ultArk - X) & " comprobaciones." & "Tiempo: " & Format((Now() - Inicio), "hh:mm:ss")
        With Ambito
            Set celda = .Find(Libro, lookAt:=xlWhole) 'Busca el Titulo entre las celdas del ambito
            'Si el titulo existe sombrea la celda, inicia un contador y sigue buscando el mismo titulo
            If Not celda Is Nothing Then
                firstAddress = celda.Address
                Do
                    With celda
                        .Font.Bold = True
                        .Font.ColorIndex = 3
                        .Interior.ColorIndex = Id_Color
                    End With
                    Cells(celda.Row, celda.Column + 2) = Cells(celda.Row, celda.Column + 2) + 1
                    Set celda = .FindNext(celda)
                Loop While Not celda Is Nothing And celda.Address <> firstAddress
            End If
        End With
    Next X

Fin = Now() 'Damos el proceso por finalizado

'Calculamos la duracion del proceso en segundos
Duracion = Hour(Fin - Inicio) * 3600 + Minute(Fin - Inicio) * 60 + Second(Fin - Inicio)

'Cuenta las celdas sombreadas en todo el ambito
For Each datax In Ambito
    If datax.Interior.ColorIndex = Id_Color Then
        CountColor = CountColor + 1
    End If
Next datax

'Muestra una ventana resumiendo el resultado del proceso
MsgBox ("El proceso ha tardado " & Format(Fin - Inicio, "hh:mm:ss") & " para un total de " & Datos & " libros." & vbCrLf & _
        "Una media de " & Format((Datos / Duracion), "0.000") & " Libros/Seg." & vbCrLf & _
        "Un total de " & CountColor & " libros ya existian en la Biblioteca, " & Datos - CountColor & " libros son nuevos.")

'Restaura la barra de estado a su origen
Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar
'Activamos los servicios desactivados
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = False

End Sub

