Option Explicit

' Definiendo las variables
Dim URL As String
Dim CantS As Integer
Dim codigo As LongLong
Dim ultimaFila As Integer
Dim i As Integer
Dim e As Integer
Dim extension As String
Dim imagenes As String
Dim contador As Integer
Dim tabla As String
Dim conglomerado As String
Dim color As String
Dim acumulado As String
Dim cuenta As Integer




Sub GeneradorImagenesVariables()
' PRODUCTO CON VARIENTE DE TALLE ===================

Sheets("Variables").Select


' Rellenando la url
URL = "https://rerda.com/imagenes/"
extension = ".jpg"
tabla = "/tabla" & extension


' Obteniendo la última fila
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count

' Bucle que recorre toda la columna
For i = 1 To ultimaFila
    
    If Cells(i, 2).Value = "variable" And Cells(i, 8) >= 1 Then
            
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        imagenes = ""
        
        For contador = 1 To CantS
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & contador & extension
        Next
                
        ' Controlando si tiene tabla
        If Cells(i, 9).Value = 1 Then
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & tabla
        End If
        
        ' Insertando el resultado completo de las imágenes
        Cells(i, 7).Value = Right(imagenes, Len(imagenes) - 1)
    End If

Next
ThisWorkbook.Save
End Sub


Sub generadorImagenesConColor()
' PRODUCTO CON VARIANTE DE COLOR ===================

Sheets("Con Color").Select

' Rellenando la url
URL = "http://localhost:8080/rerda_2/imagenes/"
extension = ".jpg"

' Obteniendo la última fila
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count

' Ordenando de mayor a menor el ID para facilitar la construcción de los enlaces del Padre.
Range("A1").Sort Key1:=Range("A1"), Order1:=xlDescending, Header:=xlNo

' Bucle que recorre toda la columna
For i = 2 To ultimaFila
    
    ' Contando cuantas veces se repite el código del producto
    If Cells(i, 4).Value = "Padre" Then
        Cells(i, 8).Value = Application.CountIf(ActiveSheet.Range("F2:F" & ultimaFila), Cells(i, 6).Value) - 1
    End If

    ' Generando los enlaces
    If Cells(i, 8).Value >= 1 Then
            
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        imagenes = ""
        
        For contador = 1 To CantS
            ' Acumula y nombra las imágenes en base primero a su código de color, seguido de su orden
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & Cells(i, 4).Value & contador & extension
        Next
        
        ' Insertando el resultado completo de las imágenes
        Cells(i, 7).Value = Right(imagenes, Len(imagenes) - 1)
        
        ' Insertando un nuevo loop para los artículos que son Padres
        If Cells(i, 4).Value = "Padre" Then
            ' Pruebo colocando el último valor de la fila anterior
            Cells(i, 7).Value = Cells(i - 1, 7).Value
        End If
    End If
    
    ' Otro loop para colocar la sumatoria concatenada de todas las imágenes de las
    ' variantes en el padre. Corroborando primero si es un Padre.
    If Cells(i, 4).Value = "Padre" Then
        acumulado = ""
        cuenta = Cells(i, 8).Value
        For e = 1 To cuenta
            acumulado = acumulado & "," & Cells(i - e, 7).Value
        Next
        
        ' Insertando el valor acumulado en la celda y eliminando la última coma.
        Cells(i, 7).Value = Right(acumulado, Len(acumulado) - 1)
    End If

Next
Range("A1").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
ThisWorkbook.Save

End Sub

Sub GeneradorImagenesSimples()
' PRODUCTO SIMPLE ===================

Sheets("Simples").Select

' Rellenando la url
URL = "http://localhost:8080/rerda_2/imagenes/"
extension = ".jpg"

' Obteniendo la última fila
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count

' Bucle que recorre toda la columna
For i = 1 To ultimaFila
    
    If Cells(i, 2).Value = "simple" And Cells(i, 8) >= 1 Then
            
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        imagenes = ""
        
        For contador = 1 To CantS
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & contador & extension
        Next
        
        ' Insertando el resultado completo de las imágenes
        Cells(i, 7).Value = Right(imagenes, Len(imagenes) - 1)
    End If

Next
ThisWorkbook.Save
End Sub
