Option Explicit
Sub GeneradorImagenes()


' Generador de enlaces de imágenes para la web ===

' Definiendo las variables
Dim URL As String
Dim CantH As Integer
Dim CantM As Integer
Dim CantS As Integer
Dim HYM As Boolean
Dim Tabla As Boolean
Dim codigo As LongLong
Dim ultimaFila As Integer
Dim i As Integer
Dim extension As String
Dim imagenes As String
Dim contador As Integer


' Rellenando la url
URL = "https://rerda.com/imagenes/"
extension = ".jpg"

' Obteniendo la última fila
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count

' Bucle que recorre toda la columna
For i = 1 To ultimaFila
    
    ' PRODUCTO SIMPLE
    If Cells(i, 2).Value = "simple" Then
        
        ' Asigando la cantidad de imágenes que tiene el producto
        CantS = Cells(i, 10).Value
        contador = 1
        imagenes = ""
        Do While contador < CantS
            If CantS = 1 Then
                imagenes = URL & Cells(i, 5).Value & "/" & contador & extension
                Exit Do
            ElseIf CantS > 1 Then
                If imagenes = "" Then
                    imagenes = URL & Cells(i, 5).Value & "/" & contador & extension
                End If
                imagenes = imagenes & "," & URL & Cells(i, 5).Value & "/" & (contador + 1) & extension
                contador = contador + 1
            End If
        Loop
        Cells(i, 6).Value = imagenes
        
    End If
    
    ' PRODUCTO CON FOTOS HOMBRE Y MUJER

    ' FOTOS CON VARIANTES DE COLOR

Next




End Sub


