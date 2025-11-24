Option Explicit

' Exporta todas las gráficas de un libro de Excel a una carpeta nueva
' creada junto al archivo. Las imágenes se guardan como PNG manteniendo
' el tamaño original de la gráfica para no perder nitidez ni proporciones
' en textos y elementos.
Sub ExportarGraficasComoImagenes()
    Dim wb As Workbook
    Dim carpetaDestino As String
    Dim rutaDestino As String
    Dim indice As Long

    Set wb = ThisWorkbook

    If Len(wb.Path) = 0 Then
        MsgBox "Guarda primero el libro para poder crear la carpeta de salida.", vbExclamation
        Exit Sub
    End If

    carpetaDestino = Format(Now, "yyyy-mm-dd_hhmmss") & "_Graficas"
    rutaDestino = wb.Path & "\" & carpetaDestino

    If Dir(rutaDestino, vbDirectory) = vbNullString Then
        MkDir rutaDestino
    End If

    indice = 1

    ' Gráficas incrustadas en hojas normales
    Dim ws As Worksheet
    Dim co As ChartObject
    For Each ws In wb.Worksheets
        For Each co In ws.ChartObjects
            indice = ExportarGrafica(co.Chart, rutaDestino, indice, ws.Name)
        Next co
    Next ws

    ' Hojas de gráfico (Chart Sheets) independientes
    Dim cs As Chart
    For Each cs In wb.Charts
        indice = ExportarGrafica(cs, rutaDestino, indice, cs.Name)
    Next cs

    MsgBox "Listo. Las gráficas se guardaron en: " & vbCrLf & rutaDestino, vbInformation
End Sub

Private Function ExportarGrafica(ByVal grafica As Chart, ByVal rutaDestino As String, _
                                 ByVal indice As Long, ByVal propietario As String) As Long
    Dim nombreArchivo As String

    nombreArchivo = rutaDestino & "\" & Format(indice, "000") & "_" & LimpiarNombre(propietario) & ".png"

    ' Exporta usando el tamaño real del gráfico para evitar distorsiones o pérdida de nitidez.
    grafica.Export Filename:=nombreArchivo, FilterName:="PNG"

    ExportarGrafica = indice + 1
End Function

Private Function LimpiarNombre(ByVal nombre As String) As String
    Dim caracteresInvalidos As Variant
    Dim i As Long

    caracteresInvalidos = Array("\\", "/", ":", "*", "?", "\"", "<", ">", "|")

    For i = LBound(caracteresInvalidos) To UBound(caracteresInvalidos)
        nombre = Replace(nombre, caracteresInvalidos(i), "_")
    Next i

    LimpiarNombre = nombre
End Function
