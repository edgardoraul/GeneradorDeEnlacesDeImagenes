Option Explicit
Sub GeneradorImagenes()


' Generador de enlaces de imágenes para la web ===

' Definiendo las variables
Dim URL As String
Dim CantS As Integer
Dim codigo As LongLong
Dim ultimaFila As Integer
Dim i As Integer
Dim extension As String
Dim imagenes As String
Dim contador As Integer
Dim tabla As String


' Rellenando la url
URL = "https://rerda.com/imagenes/"
extension = ".jpg"
tabla = "/tabla" & extension

' Obteniendo la última fila
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count

' Bucle que recorre toda la columna
For i = 1 To ultimaFila
    
    ' PRODUCTO SIMPLE
    If Cells(i, 2).Value = "simple" Or Cells(i, 2).Value = "variable" Then
        
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        contador = 1
        imagenes = ""
        Do While contador <= CantS
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & contador & extension
            contador = contador + 1
        Loop
                
        ' Controlando si tiene tabla
        If Cells(i, 9).Value = 1 Then
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & tabla
        End If
        
        ' Insertando el resultado completo de las imágenes
        Cells(i, 7).Value = Right(imagenes, Len(imagenes) - 1)
    
    
    ' PRODUCTO VARIANTES DE TALLE
    ElseIf Cells(i, 2).Value = "variation" And Cells(i, 4).Value = 1 Then
        
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 8).Value
        contador = 1
        imagenes = ""
        Do While contador <= CantS
            imagenes = imagenes & "," & URL & Cells(i, 6).Value & "/" & contador & extension
            contador = contador + 1
        Loop
        Cells(i, 6).Value = Right(imagenes, Len(imagenes) - 1)
    End If

Next




End Sub
