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
Dim tipo As String
Dim ruta As String
Dim rutaImgRenombradas As String
Dim subCarpeta As String
Dim origen As String
Dim archivoNuevo As String
Dim cantidadImg As Integer
Dim destino As String
Dim archivoAntiguo As String
Dim xPath As String
Dim xFile As String
Dim xCount As Integer
Dim cantidad As String


' Una función para obtener las rutas de carpetas
Function OBTENER_RUTA_CARPETA_ARCHIVO(ruta As String) As String
    Set objeto = New FileSystemObject
    Set Archivo = objeto.GetFile(ruta)
    OBTENER_RUTA_CARPETA_ARCHIVO = Archivo.ParentFolder.Path
End Function


Sub GeneradorImagenesVariables()
' PRODUCTO CON VARIENTE DE TALLE ===================

Sheets("Variables").Activate


' Rellenando la url
' URL = "http://localhost:8080/rerda_2/imagenes/"
URL = Sheets("Constantes").Range("B1").Value
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

Sheets("Con Color").Activate

' Rellenando la url
URL = Sheets("Constantes").Range("B1").Value
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

Sheets("Simples").Activate

' Rellenando la url
URL = Sheets("Constantes").Range("B1").Value
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


Sub ElegirCarpeta()
    ' Aplicación que sirve para elegir la carpeta de las imágenes

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.Path & "\"
        .Title = "Seleccionar carpeta"
        .Show
    
        If .SelectedItems.Count = 0 Then
            MsgBox "Nada"
        Else
            ruta = .SelectedItems(1)
            MsgBox ruta
        End If
        copiarImgVariables
    End With
End Sub

Sub copiarImgVariables()
    ' DESCRIPCION: Copia y renombra imágenes con variantes de talles
    
    ' Creamos coordenadas para trabajar
    Range("A1").Select
    ultimaFila = Range(Selection, Selection.End(xlDown)).Count
    Dim subCarpeta As String
    Dim Carpeta As String
    Dim archivoViejo As String
    Dim archivoNuevo As String
    Dim origen As String
    Dim destino As String
    Dim fs As Object
    Dim cantidadImg As Integer
    Dim skuActual As String
    Dim skuAnterior As String
    Dim codigoActual As String
    Dim codigoAnterior As String
    Dim Seguir As String
    
    extension = ".jpg"
    ruta = "D:\xampp\htdocs\rerda_2\OneDrive\imagenes\"
    rutaImgRenombradas = ruta & "..\Dragonfish Color y Talle\Articulos\"
    

    ' Bucle para recorrer toda la columna de los códigos y todas las carpetas con las imágenes
    For i = 2 To ultimaFila
        ' Posicionándose en lo que importa
        Cells(i, 1).Activate
        
        ' Código -> Corresponde a la carpeta en la que están las imágenes numeradas
        subCarpeta = Cells(i, 6).Value
        
        
        ' Contar la cantidad de imágenes que hay una carpeta determinada
        xPath = ruta & subCarpeta & "\*" & extension
        xFile = Dir(xPath)
        
        xCount = 0
        Do While xFile <> ""
            xCount = xCount + 1
            If xFile = "tabla.jpg" Then
                xCount = xCount - 1
                Cells(i, 9).Value = 1
            End If
            xFile = Dir()
        Loop
        
        ' Insertando el resultado encontrado en la planilla
        'If Cells(i, 9).Value = 1 Then
         '   xCount = xCount - 1
        'End If
        Cells(i, 8).Value = xCount
        
        ' Extrayendo de la planilla la cantidad de imágenes
        cantidadImg = xCount
        
        Debug.Print "El Código " & subCarpeta & " tiene " & cantidadImg & " imágenes."
        
        If cantidadImg < 1 Then
            Cells(i, 10).Value = "El código " & subCarpeta & " no tiene imágenes."
            Cells(i, 8).Value = ""
            GoTo Seguir
        Else
            Cells(i, 10).Value = ""
        End If
        
        Debug.Print "Estamos en la fila N° " & i
        
        
        
        
        
        ' SKU limpio. Corresponde al código en si mismo que tiene el producto
        skuActual = Left(Cells(i, 3).Value, 7)
        codigoActual = Cells(i, 6).Value
        
        ' Controlando que a partir del segundo item real vaya este control
        If i > 2 Then
            skuAnterior = Left(Cells(i - 1, 3).Value, 7)
            codigoAnterior = Cells(i - 1, 6).Value
        End If
        
        ' Controlando si sku actual tiene el mismo código que el sku anterior
        If skuActual = codigoActual Then
            'Debug.Print "El sku actual " & skuActual & " coincide con el código " & codigoActual
        ElseIf skuActual <> codigoActual And codigoActual = codigoAnterior Then
            'Debug.Print "El sku actual " & skuActual & " es talle grande del código " & codigoAnterior
        End If
        
        
        
        ' Controlando si tiene tabla de talles
        If Cells(i, 9).Value = 1 Then
            cantidadImg = cantidadImg + 1
        End If
        
        
        ' Creando nuevos nombres de archivos de fotos mediante bucle
        For e = 1 To cantidadImg
            
            ' Carpeta y nombre de archivo de Origen
            If e = cantidadImg And Cells(i, 9).Value = 1 Then
                origen = ruta & subCarpeta & "\" & "tabla" & extension
            Else
                origen = ruta & subCarpeta & "\" & e & extension
            End If
            
            ' Nuevo nombre de archivo
            archivoNuevo = skuActual & "'''" & e & extension
            
            ' Carpeta y nombre nuevo de destino
            destino = rutaImgRenombradas & archivoNuevo
            FileCopy origen, destino
            Debug.Print origen & " está copiado como " & destino
            
        Next
Seguir:
    Next
    
End Sub
Sub copiarImgColor()
' DESCRIPCION: Copia y renombra imágenes con variantes de COLOR
' Creamos coordenadas para trabajar
Range("A1").Select
ultimaFila = Range(Selection, Selection.End(xlDown)).Count
extension = ".jpg"
ruta = "D:\xampp\htdocs\rerda_2\OneDrive\imagenes\"
rutaImgRenombradas = ruta & "..\Dragonfish Color y Talle\Articulos\"



' Copiando Imágenes. Recorremos toda la tabla desde arriba hasta abajo
For i = 2 To ultimaFila
    'Definiendo la cantidad de imágenes que tiene esta variante
    codigo = Cells(i, 6).Value
    xPath = ruta & codigo & "\*" & extension
    xFile = Dir(xPath)
    If xFile = "" Then
        Cells(i, 8).Value = "Sin imágenes"
        GoTo Seguir
    End If
    
    'Averiguando cuántas imágenes hay en la variante o padre seleccionada
    xCount = 0
    Do While xFile <> ""
        xCount = xCount + 1
        Cells(i, (8 + xCount)).Value = xFile
        ' Aquí ya cambia de valor
        xFile = Dir()
    Loop
    
    
    'Extrayendo la cantidad de imágenes que tiene cada publicación
    cantidadImg = xCount
    
    'Anotando el resultado
    Cells(i, 8).Value = cantidadImg
    
    'Renombrando cada imagen y copiándola al destino
    For e = 1 To cantidadImg

        ' Foto de portada. Una sola.
        If Cells(i, (8 + e)).Value = "1.jpg" Then
            archivoAntiguo = Cells(i, (8 + e)).Value
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "'''" & extension
        
        ElseIf Cells(i, (8 + e)).Value = "2.jpg" Then
            archivoAntiguo = Cells(i, (8 + e)).Value
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "'''1" & extension
        
        Else
            ' Fotos de las variantes de color
            color = Left(Cells(i, (8 + e)).Value, 2)
            cantidad = Mid(Cells(i, (8 + e)).Value, 3, (Len(Cells(i, (8 + e))) - Len(extension) - 2))
            archivoAntiguo = color & cantidad & extension
            origen = ruta & codigo & "\" & archivoAntiguo
            archivoNuevo = codigo & "'" & color & "''" & cantidad & extension
        End If
        
        'Definiendo el destino final del archivo de la imagen
        destino = rutaImgRenombradas & archivoNuevo
        
        'Copiando el achivo con el nuevo nombre
        Debug.Print "Fila " & i & " tiene " & cantidadImg & " # " & archivoAntiguo & " -> " & archivoNuevo
               
        FileCopy origen, destino
        
        Debug.Print origen & " -> " & destino
Seguir:
    Next
    
Next
    
End Sub