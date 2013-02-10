package cager.parser.test;

public class main{

	 void AltaEst_Click ( ) {

/* 
 Private Sub AltaEst_Click()
    progreso.Value = 0
    Form1.Show
End Sub 
 */
	}

	 void BajaEst_Click ( ) {

/* 
 Private Sub BajaEst_Click()
    progreso.Value = 0
    Form2.Show
End Sub 
 */
	}

	 void fechaAux_Change ( ) {

/* 
 












Private Sub fechaAux_Change()
    calendar.Value = fechaAux.Text
End Sub 
 */
	}

	 void Form_Initialize ( ) {

/* 
 

Private Sub Form_Initialize()
    AFOROS.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BaseDatos.mdb;Persist Security Info=False"
    AFOROS.RecordSource = "ESTDATAUX"
End Sub 
 */
	}

	 void Form_Load ( ) {

/* 
 


Private Sub Form_Load()
    AFOROS.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BaseDatos.mdb;Persist Security Info=False"
    AFOROS.RecordSource = "ESTDATAUX"
    
    restab

    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 8700)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 9700)
    
salir:
    Exit Sub
err_Load:
    cnn.Close
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al cargar la aplicación")
    Resume salir
End Sub 
 */
	}

	 void Form_Unload ( Integer Cancel ) {

/* 
 Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub 
 */
	}

	 void abrirBD_Click ( ) {

/* 
 Private Sub abrirBD_Click()
    progreso.Value = 0
    Shell "Explorer.exe " & App.Path & "\BaseDatos.mdb", vbMaximizedFocus
End Sub 
 */
	}

	 void abrir_Click ( ) {

/* 
 Private Sub abrir_Click()
    abrirFichero
End Sub 
 */
	}

	 void abrirFichero ( ) {

/* 
 

Private Sub abrirFichero()
On Error GoTo ErrorProducido
    Dim sFile As String
    
    progreso.Value = 0
    
    With dlgCommonDialog
        .DialogTitle = "Abrir"
        .CancelError = False
        .Filter = "Archivos FIM|*.FIM"
        .InitDir = App.Path
        .ShowOpen
        
        sName = .FileTitle
        
        If Len(.fileName) = 0 Then
            Exit Sub
        End If
        sFile = .fileName
    End With
    
    ruta.Text = sFile
    sPath = Left(sFile, (Len(sFile) - Len(sName)) - 1)
    
    
salir:
    Exit Sub
ErrorProducido:
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al abrir el archivo")
    Resume salir
End Sub 
 */
	}

	 void imprimir_Click ( ) {

/* 
 Private Sub imprimir_Click()
    On Error GoTo ErrorProducido
    progreso.Value = 0
    With CommonDialog1
        .DialogTitle = "Imprimir"
        .CancelError = False
        .fileName = App.Path & "/BaseDatos.mdb"
        .PrinterDefault = False
        .Filter = "Bases de Datos Access (*.mdb)|*.mdb"
        .ShowPrinter
    End With
salir:
    Exit Sub
ErrorProducido:
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al imprimir")
    Resume salir

End Sub 
 */
	}

	 void importar_Click ( ) {

/* 
 
Private Sub importar_Click()
    
    On Error GoTo Err_importar_Click
    
    If ruta.Text = "" Or ruta.Text = " " Then
        a = MsgBox("Antes, debe seleccionar un fichero", vbExclamation, "Error de Importación")
        progreso.Value = 0
        abrirFichero
        Exit Sub
    End If
    'Preguntamos si quiere importar todos los del directorio
    todos = MsgBox("¿Desea procesar todos los archivos *.FIM de este directorio?", vbOKCancel, "Procesar Todos")
    If todos = 2 Then
        importarArchivo (ruta.Text)
    Else
        'LISTAMOS LOS FICHEROS de la ruta
        Dim fso As Object 'New FileSystemObject
        Dim carpeta As Object 'Folder
        Dim archivos As Object 'Files
        Dim archivo As Object 'File
        Dim fileName, file, fileExt As String
    
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set carpeta = fso.GetFolder(sPath)
        Set archivos = carpeta.Files
    
        'importamos cada archivo (siempre y cuando sea *.fim)
        For Each archivo In archivos
            fileName = archivo.Name
            sName = fileName
            fileExt = Mid(fileName, Len(fileName) - 2, 3)
            If fileExt = "fim" Or fileExt = "FIM" Or fileExt = "Fim" Then
                file = sPath & "\" & fileName
                importarArchivo (file)
            End If
        Next
        
        Set archivos = Nothing
        Set carpeta = Nothing
        Set fso = Nothing
    End If
    ruta = ""
Exit_importar_Click:
        Exit Sub
Err_importar_Click:
        a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error de Importación")
        ruta = ""
        Resume Exit_importar_Click
End Sub 
 */
	}

	 void importarArchivo ( String file ) {

/* 
 Private Sub importarArchivo(file As String)
    
    'Variables de base de datos
    Dim cnn As ADODB.Connection
    Dim regAux As ADODB.Recordset
    Dim baseC As Database
    Set baseC = OpenDatabase(App.Path & "\BaseDatos.mdb")
    
    'Variables a nivel de fichero
    Dim numBloques As Integer
    Dim numLineasBloque As Integer
    Dim totalLineas As Integer
    Dim linea As String
    Dim cont As Integer
    Dim QTencontrado As Boolean
    Dim QLencontrado As Boolean
    QTencontrado = False
    QLencontrado = False
    
    'Variables a nivel de linea
    Dim idEstacion As String
    Dim fecha As String
    'fecha = "12/28/06"
    Dim fechaAux As Date
    Dim horaInicio As Integer
    Dim hora As Integer
    Dim intervalo As Integer
    Dim canal As Integer
    Dim totVeh As Integer
    Dim velMed As Integer
    Dim totPes As Integer
    Dim limitePeso As String
    
    'Variables auxiliares
    Dim ini As Integer
    Dim aux As String
    Dim acumulado As Integer
    
    
    
    
    On Error GoTo Err_importarArchivo
              
        
        'abrimos el fichero
        Open file For Input Shared As #2
        
        'obtenemos el nombre de la estacion a partir del nombre del fichero
        Dim s As String
        For i = 1 To Len(sName)
            s = Mid(sName, i, 1)
            If s = "_" Then
                idEstacion = Mid(sName, 1, i - 1)
                Exit For
            End If
        Next
        
        
        
        ' Establezco la conexión con la base de datos actual
        Set cnn = New ADODB.Connection
        With cnn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\BaseDatos.mdb"
            .Open
            
            'Comprobamos que la estacion existe en nuestra base de datos
            consulta = "SELECT IDENTIFICACION FROM EST WHERE IDENTIFICACION='" & idEstacion & "';"
            Set regAux = .Execute(consulta)
            If regAux.EOF <> False Then
                a = MsgBox("La estación a la que hace referencia no está registrada" & vbCrLf & "¿Desea dar de alta a la estación '" & idEstacion & "' en este momento?", vbOKCancel, "Error de Importación")
                If a = 1 Then
                    Form1.Show
                    Form1.identificacionT.Text = idEstacion
                    cnn.Close
                    Close #2
                    Exit Sub
                Else
                    cnn.Close
                    Close #2
                    Exit Sub
                End If
            End If
            
            'Extraemos la primera linea para obtener informacion necesaria para recorrer el fichero
            Line Input #2, linea
            longitud = Len(linea)
            If (longitud <> 78) Then
                a = MsgBox("El fichero *.FIM no es correcto", vbOKCancel, "Error de Importación")
                Close #2
                Exit Sub
            End If
            
            cont = 1
            'MsgBox linea
            tipo = Mid(linea, 58, 2)
            'MsgBox tipo
            numBloques = Mid(linea, 76, 4)
            'MsgBox numBloques
            numLineasBloque = Mid(linea, 71, 4)
            'MsgBox numLineasBloque
            totalLineas = numBloques * numLineasBloque + numBloques
            'MsgBox totalLineas
            progreso.Max = totalLineas
            
            
            Do
                
                
                
                
                'Estos son los tipos de bloques que tenemos en cuenta
                If tipo = "QT" Or tipo = "QP" Or tipo = "VT" Or tipo = "QL" Or tipo = "LC" Then
                    
                    'Calculamos la clave primaria
                    fecha = Mid(linea, 38, 2) & "/" & Mid(linea, 33, 2) & "/" & Mid(linea, 28, 2)
                    'MsgBox fecha
                    horaInicio = Mid(linea, 43, 2)
                    'MsgBox horaInicio
                    canal = Mid(linea, 69, 1) + 1
                    'MsgBox canal
                    intervalo = Mid(linea, 51, 4)
                    'MsgBox intervalo
                
                
                    'Insertamos en la base de datos según el tipo del bloque
                    
                    
                    
                    'Conteo de vehiculos
                    If tipo = "QT" Then
                        QTencontrado = True
                        hora = horaInicio
                        acumulado = -1
                        
                        For i = 1 To numLineasBloque
                            Line Input #2, linea
                            'MsgBox linea
                            longitud = Len(linea)
                            If (longitud <> 78) Then
                                a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                progreso.Value = 0
                                Close #2
                                Exit Sub
                            End If
                            cont = cont + 1
                            ini = 1
                            For j = 1 To 12
                                aux = Mid(linea, ini, 5)
                                If aux <> "     " Then
                                    If intervalo = 60 Then
                                        totVeh = aux
                                        ini = ini + 6
                                        '##################
                                        consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, TOTVEH FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                        Set regAux = .Execute(consulta)
                                        If regAux.EOF = False Then
                                            If ("" & regAux![totVeh]) = "" Then
                                                'MsgBox "hay fila y campo ES NULO"
                                                'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Else
                                                'MsgBox "hay fila pero campo NO ES NULO"
                                                'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            End If
                                        Else
                                            'MsgBox "No hay fila"
                                            'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                            .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, TOTVEH) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totVeh & ");"
                                        End If
                                        '##################
                                        hora = hora + 1
                                        If hora = 24 Then
                                            hora = 0
                                            fecha = incrementaFecha(fecha)
                                        End If
                                    End If
                                    If intervalo = 30 Then
                                        If acumulado = -1 Then
                                            acumulado = aux
                                            ini = ini + 6
                                        Else
                                            totVeh = acumulado + aux
                                            ini = ini + 6
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, TOTVEH FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![totVeh]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, TOTVEH) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totVeh & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                            acumulado = -1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    
                    
                    
                    
                    
                    'Conteo de vehiculos
                    If tipo = "QP" And QTencontrado = False Then
                        hora = horaInicio
                        acumulado = -1
                        
                        For i = 1 To numLineasBloque
                            Line Input #2, linea
                            'MsgBox linea
                            longitud = Len(linea)
                            If (longitud <> 78) Then
                                a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                progreso.Value = 0
                                Close #2
                                Exit Sub
                            End If
                            cont = cont + 1
                            ini = 1
                            For j = 1 To 12
                                aux = Mid(linea, ini, 5)
                                If aux <> "     " Then
                                    If intervalo = 60 Then
                                        totVeh = aux
                                        ini = ini + 6
                                        '##################
                                        consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, TOTVEH FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                        Set regAux = .Execute(consulta)
                                        If regAux.EOF = False Then
                                            If ("" & regAux![totVeh]) = "" Then
                                                'MsgBox "hay fila y campo ES NULO"
                                                'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Else
                                                'MsgBox "hay fila pero campo NO ES NULO"
                                                'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            End If
                                        Else
                                            'MsgBox "No hay fila"
                                            'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                            .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, TOTVEH) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totVeh & ");"
                                        End If
                                        '##################
                                        hora = hora + 1
                                        If hora = 24 Then
                                            hora = 0
                                            fecha = incrementaFecha(fecha)
                                        End If
                                    End If
                                    If intervalo = 30 Then
                                        If acumulado = -1 Then
                                            acumulado = aux
                                            ini = ini + 6
                                        Else
                                            totVeh = acumulado + aux
                                            ini = ini + 6
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, TOTVEH FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![totVeh]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET TOTVEH=" & totVeh & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, TOTVEH) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totVeh & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                            acumulado = -1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    
                    
                    
                    
                    
                    'Conteo de la velocidad de los vehiculos
                    If tipo = "VT" Then
                        hora = horaInicio
                        acumulado = -1
                        
                        For i = 1 To numLineasBloque
                            Line Input #2, linea
                            'MsgBox linea
                            longitud = Len(linea)
                            If (longitud <> 78) Then
                                a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                progreso.Value = 0
                                Close #2
                                Exit Sub
                            End If
                            cont = cont + 1
                            ini = 1
                            For j = 1 To 12
                                aux = Mid(linea, ini, 3)
                                If aux <> "   " Then
                                    If intervalo = 60 Then
                                        velMed = aux
                                        ini = ini + 4
                                        '##################
                                        consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VELMED FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                        Set regAux = .Execute(consulta)
                                        If regAux.EOF = False Then
                                            If ("" & regAux![velMed]) = "" Then
                                                'MsgBox "hay fila y campo ES NULO"
                                                'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                .Execute "UPDATE ESTDAT SET VELMED=" & velMed & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Else
                                                'MsgBox "hay fila pero campo NO ES NULO"
                                                'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                .Execute "UPDATE ESTDAT SET VELMED=" & velMed & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            End If
                                        Else
                                            'MsgBox "No hay fila"
                                            'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                            .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VELMED) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & velMed & ");"
                                        End If
                                        '##################
                                        hora = hora + 1
                                        If hora = 24 Then
                                            hora = 0
                                            fecha = incrementaFecha(fecha)
                                        End If
                                    End If
                                    If intervalo = 30 Then
                                        If acumulado = -1 Then
                                            acumulado = aux
                                            ini = ini + 4
                                        Else
                                            velMed = acumulado + aux
                                            ini = ini + 4
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VELMED FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![velMed]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET VELMED=" & velMed & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET VELMED=" & velMed & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VELMED) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & velMed & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                            acumulado = -1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    
                    
                    
                    
                    'Conteo de vehiculos pesados
                    If tipo = "QL" Then
                        QLencontrado = True
                        hora = horaInicio
                        acumulado = -1
                        
                        For i = 1 To numLineasBloque
                            Line Input #2, linea
                            'MsgBox linea
                            longitud = Len(linea)
                            If (longitud <> 78) Then
                                a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                progreso.Value = 0
                                Close #2
                                Exit Sub
                            End If
                            cont = cont + 1
                            ini = 1
                            For j = 1 To 12
                                aux = Mid(linea, ini, 5)
                                If aux <> "     " Then
                                    If intervalo = 60 Then
                                        totPes = aux
                                        ini = ini + 6
                                        '##################
                                        consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VEHPES FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                        Set regAux = .Execute(consulta)
                                        If regAux.EOF = False Then
                                            If ("" & regAux![VEHPES]) = "" Then
                                                'MsgBox "hay fila y campo ES NULO"
                                                'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Else
                                                'MsgBox "hay fila pero campo NO ES NULO"
                                                'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            End If
                                        Else
                                            'MsgBox "No hay fila"
                                            'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                            .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VEHPES) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totPes & ");"
                                        End If
                                        '##################
                                        hora = hora + 1
                                        If hora = 24 Then
                                            hora = 0
                                            fecha = incrementaFecha(fecha)
                                        End If
                                    End If
                                    If intervalo = 30 Then
                                        If acumulado = -1 Then
                                            acumulado = aux
                                            ini = ini + 6
                                        Else
                                            totVeh = acumulado + aux
                                            ini = ini + 6
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VEHPES FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![VEHPES]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VEHPES) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totPes & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                            acumulado = -1
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                    
                    
                    
                    
                    
                    'Conteo de vehiculos pesados
                    If tipo = "LC" And QLencontrado = False Then
                        limitePeso = Mid(linea, 1, 2)
                        If limitePeso >= 37 Then
                            hora = horaInicio
                            acumulado = -1
                            
                            For i = 1 To numLineasBloque
                                Line Input #2, linea
                                'MsgBox linea
                                longitud = Len(linea)
                                If (longitud <> 78) Then
                                    a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                                    progreso.Value = 0
                                    Close #2
                                    Exit Sub
                                End If
                                cont = cont + 1
                                ini = 1
                                For j = 1 To 12
                                    aux = Mid(linea, ini, 5)
                                    If aux <> "     " Then
                                        If intervalo = 60 Then
                                            totPes = aux
                                            ini = ini + 6
                                            '##################
                                            consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VEHPES FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                            Set regAux = .Execute(consulta)
                                            If regAux.EOF = False Then
                                                If ("" & regAux![VEHPES]) = "" Then
                                                    'MsgBox "hay fila y campo ES NULO"
                                                    'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                    .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Else
                                                    'MsgBox "hay fila pero campo NO ES NULO"
                                                    'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                    .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                End If
                                            Else
                                                'MsgBox "No hay fila"
                                                'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VEHPES) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totPes & ");"
                                            End If
                                            '##################
                                            hora = hora + 1
                                            If hora = 24 Then
                                                hora = 0
                                                fecha = incrementaFecha(fecha)
                                            End If
                                        End If
                                        If intervalo = 30 Then
                                            If acumulado = -1 Then
                                                acumulado = aux
                                                ini = ini + 6
                                            Else
                                                totVeh = acumulado + aux
                                                ini = ini + 6
                                                '##################
                                                consulta = "SELECT IDEST, FECHAMED, HORAMED, CANAL, VEHPES FROM ESTDAT WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                Set regAux = .Execute(consulta)
                                                If regAux.EOF = False Then
                                                    If ("" & regAux![VEHPES]) = "" Then
                                                        'MsgBox "hay fila y campo ES NULO"
                                                        'ACTUALIZAMOS PONIENDO EL CAMPO NUEVO
                                                        .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                    Else
                                                        'MsgBox "hay fila pero campo NO ES NULO"
                                                        'ACTUALIZAMOS INCREMENTANDO EL CAMPO
                                                        .Execute "UPDATE ESTDAT SET VEHPES=" & totPes & " WHERE IDEST='" & idEstacion & "' AND FECHAMED='" & fecha & "' AND HORAMED=" & hora & " AND canal=" & canal & ";"
                                                    End If
                                                Else
                                                    'MsgBox "No hay fila"
                                                    'INSERTAMOS EL CAMPO Y TODOS LOS DEMÁS DATOS
                                                    .Execute "INSERT INTO ESTDAT (IDEST, FECHAMED, HORAMED, CANAL, VEHPES) VALUES ('" & idEstacion & "', '" & fecha & "', " & hora & ", " & canal & ", " & totPes & ");"
                                                End If
                                                '##################
                                                hora = hora + 1
                                                If hora = 24 Then
                                                    hora = 0
                                                    fecha = incrementaFecha(fecha)
                                                End If
                                                acumulado = -1
                                            End If
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    End If
                    
                    

                    
               
                Else
                    'SE IGNORAN LOS DEMÁS BLOQUES
                
                End If
                

                'cogemos la siguiente linea
                If cont < totalLineas Then
                    Line Input #2, linea
                    longitud = Len(linea)
                    If (longitud <> 78) Then
                        a = MsgBox("El fichero *.FIM no es correcto", vbOKOnly, "Error de Importación")
                        progreso.Value = 0
                        Close #2
                        Exit Sub
                    End If
                    tipo = Mid(linea, 58, 2)
                End If
                progreso = cont
                cont = cont + 1
                                  
                
            Loop While cont <= totalLineas

            
            .Close
            Close #2
            
        End With
        a = MsgBox("Los datos de '" & sName & "' se importaron correctamente", vbOKOnly, "Importación Correcta")
        
        restab
        
Exit_importarArchivo:
        Exit Sub
Err_importarArchivo:
        a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error de Importación")
        progreso.Value = 0
        cnn.Close
        Close #2
        Resume Exit_importarArchivo
End Sub 
 */
	}

	 void incrementaFecha ( String fecha ) {

/* 
 Private Function incrementaFecha(ByVal fecha As String) As String
    dia = Mid(fecha, 1, 2)
    mes = Mid(fecha, 4, 2)
    ano = Mid(fecha, 7, 2)

    If (ano Mod 4 = 0) Then
        'bisiesto
        If (mes = 1 Or mes = 3 Or mes = 5 Or mes = 7 Or mes = 8 Or mes = 10 Or mes = 12) Then
            'mes de 31 dias
            If (dia = 31) Then
                dia = 1
                If (mes = 12) Then
                    mes = 1
                    ano = ano + 1
                Else
                    mes = mes + 1
                End If
            Else
                dia = dia + 1
            End If
        Else
            If (mes = 2) Then
                'febrero de 29 dias
                If (dia = 29) Then
                    dia = 1
                    mes = mes + 1
                Else
                    dia = dia + 1
                End If
            Else
                'resto meses 30 dias
                If (dia = 30) Then
                    dia = 1
                    mes = mes + 1
                Else
                    dia = dia + 1
                End If
            End If
        End If
    Else
        'No bisiesto
        If (mes = 1 Or mes = 3 Or mes = 5 Or mes = 7 Or mes = 8 Or mes = 10 Or mes = 12) Then
            'mes de 31 dias
            If (dia = 31) Then
                dia = 1
                If (mes = 12) Then
                    mes = 1
                    ano = ano + 1
                Else
                    mes = mes + 1
                End If
            Else
                dia = dia + 1
            End If
        Else
            If (mes = 2) Then
                'febrero de 28 dias
                If (dia = 28) Then
                    dia = 1
                    mes = mes + 1
                Else
                    dia = dia + 1
                End If
            Else
                'resto meses 30 dias
                If (dia = 30) Then
                    dia = 1
                    mes = mes + 1
                Else
                    dia = dia + 1
                End If
            End If
        End If
    End If

    If (Len(dia) = 1) Then
        dia = "0" & dia
    End If
    If (Len(mes) = 1) Then
        mes = "0" & mes
    End If
    If (Len(ano) = 1) Then
        ano = "0" & ano
    End If
    

    incrementaFecha = dia & "/" & mes & "/" & ano
    
    
End Function 
 */
	}

	 void buscar_Click ( ) {

/* 
 Private Sub buscar_Click()
    Dim cnn As ADODB.Connection
    Dim SQL As String
    
    progreso.Value = 0
    
    On Error GoTo Err_buscar_Click
                      
        
        ' Establezco la conexión con la base de datos actual
        Set cnn = New ADODB.Connection
        With cnn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\BaseDatos.mdb"
            .Open
            
            'muestro los datos de la estacion
            consulta = "SELECT IDENTIFICACION FROM EST WHERE IDENTIFICACION='" & estText.Text & "';"
            Set regAux = .Execute(consulta)
            If regAux.EOF <> False Then
                a = MsgBox("La estación a la que hace referencia no está registrada", vbExclamation, "Error de Filtrado")
                restab
                estText.Text = ""
                estText.SetFocus
                Exit Sub
            Else
                consulta = "SELECT IDENTIFICACION, TIPOESTACION, NUMCANALES, CARRETERA, PK, DELEGACION, SITUACION, CANAL1, CANAL2, CANAL3, CANAL4 FROM EST WHERE IDENTIFICACION='" & estText.Text & "';"
                Set regAux = .Execute(consulta)
                estText.Text = "" & regAux![identificacion]
                loc.Visible = True
                locText.Visible = True
                locText.Text = "" & regAux![carretera]
                pk.Visible = True
                pkText.Visible = True
                pkText.Text = "" & regAux![pk]
                tipol.Visible = True
                tipoText.Visible = True
                tipoText.Text = "" & regAux![tipoestacion]
                del.Visible = True
                delText.Visible = True
                delText.Text = "" & regAux![delegacion]
                tra.Visible = True
                traText.Visible = True
                traText.Text = "" & regAux![situacion]
                c1.Visible = True
                c1Text.Visible = True
                c1Text.Text = "" & regAux![canal1]
                c2.Visible = True
                c2Text.Visible = True
                c2Text.Text = "" & regAux![canal2]
                c3.Visible = True
                c3Text.Visible = True
                c3Text.Text = "" & regAux![canal3]
                c4.Visible = True
                c4Text.Visible = True
                c4Text.Text = "" & regAux![canal4]
            End If
            
            

            
            'realizo el filtrado
            .Execute "DELETE * FROM ESTDATAUX"
            If estText.Text = "" Or estText.Text = " " Then
                a = MsgBox("Debe introducir el nombre de una estación", vbExclamation, "Error de Filtrado")
            Else
                .Execute "INSERT INTO ESTDATAUX " & _
                " (IDEST, FECHAMED, HORAMED, TOTVEH1, VELMED1, VEHPES1, TOTVEH2, VELMED2, VEHPES2, TOTVEHT, VELMEDT, VEHPEST) " & _
                " SELECT A.IDEST, A.FECHAMED, A.HORAMED, A.TOTVEH, A.VELMED, A.VEHPES, B.TOTVEH, B.VELMED, B.VEHPES, A.TOTVEH+B.TOTVEH, ((A.VELMED+B.VELMED)/2), A.VEHPES+B.VEHPES " & _
                " FROM ESTDAT AS A, ESTDAT AS B " & _
                " WHERE A.IDEST='" & estText.Text & "' AND A.IDEST= B.IDEST AND A.FECHAMED=B.FECHAMED AND A.HORAMED=B.HORAMED AND A.CANAL=1 AND B.CANAL=2;"
            End If
            AFOROS.RecordSource = "ESTDATAUX"
            a = MsgBox("Los datos se filtrarán", vbOKOnly, "Filtrado")
            'estText.Text = ""
            calendar.Year = Year(Date)
            calendar.Month = Month(Date)
            calendar.Day = Day(Date)
            calendar.Value = fechaAux.Text
            AFOROS.Refresh
            ' Cerramos la conexión
            .Close
            
        End With
        
        'calendar.Value = Today
        
Exit_buscar_Click:
        Exit Sub
Err_buscar_Click:
        a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error de Filtrado")
        cnn.Close
        Resume Exit_buscar_Click
End Sub 
 */
	}

	 void calendar_Click ( ) {

/* 
 Private Sub calendar_Click()
    Dim cnn As ADODB.Connection
    Dim SQL As String
    Dim fecha As String
    fecha = calendar.Value
    
    dia = Mid(fecha, 1, 2)
    mes = Mid(fecha, 4, 2)
    ano = Mid(fecha, 9, 2)
    
    fecha = dia & "/" & mes & "/" & ano
    'MsgBox fecha
    
    progreso.Value = 0
    
    On Error GoTo Err_calendar_Click
        'construimos la sentencia SQL
        If estText.Text = "" Or estText.Text = " " Then
            SQL = "INSERT INTO ESTDATAUX " & _
                " (IDEST, FECHAMED, HORAMED, TOTVEH1, VELMED1, VEHPES1, TOTVEH2, VELMED2, VEHPES2, TOTVEHT, VELMEDT, VEHPEST) " & _
                " SELECT A.IDEST, A.FECHAMED, A.HORAMED, A.TOTVEH, A.VELMED, A.VEHPES, B.TOTVEH, B.VELMED, B.VEHPES, A.TOTVEH+B.TOTVEH, A.VELMED+B.VELMED, A.VEHPES+B.VEHPES " & _
                " FROM ESTDAT AS A, ESTDAT AS B " & _
                " WHERE A.FECHAMED = '" & fecha & "' AND A.IDEST= B.IDEST AND A.FECHAMED=B.FECHAMED AND A.HORAMED=B.HORAMED AND A.CANAL=1 AND B.CANAL=2;"
        Else
            SQL = "INSERT INTO ESTDATAUX " & _
                " (IDEST, FECHAMED, HORAMED, TOTVEH1, VELMED1, VEHPES1, TOTVEH2, VELMED2, VEHPES2, TOTVEHT, VELMEDT, VEHPEST) " & _
                " SELECT A.IDEST, A.FECHAMED, A.HORAMED, A.TOTVEH, A.VELMED, A.VEHPES, B.TOTVEH, B.VELMED, B.VEHPES, A.TOTVEH+B.TOTVEH, A.VELMED+B.VELMED, A.VEHPES+B.VEHPES " & _
                " FROM ESTDAT AS A, ESTDAT AS B " & _
                " WHERE A.FECHAMED = '" & fecha & "' AND A.IDEST = '" & estText.Text & "' AND A.IDEST = B.IDEST AND A.FECHAMED=B.FECHAMED AND A.HORAMED=B.HORAMED AND A.CANAL=1 AND B.CANAL=2;"
        End If

        ' Establezco la conexión con la base de datos actual
        Set cnn = New ADODB.Connection
        With cnn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\BaseDatos.mdb"
            .Open
            .Execute "DELETE * FROM ESTDATAUX;"
            .Execute SQL
            AFOROS.RecordSource = "ESTDATAUX"
            a = MsgBox("Los datos se filtrarán", vbOKOnly, "Filtrado")
            AFOROS.Refresh
            ' Cerramos la conexión
            .Close
        End With
        
Exit_calendar_Click:
        Exit Sub
Err_calendar_Click:
        cnn.Close
        a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error de Filtrado")
        Resume Exit_calendar_Click
End Sub 
 */
	}

	 void info_Click ( ) {

/* 
 Private Sub info_Click()
    progreso.Value = 0
    examinarArchivo (App.Path & "/BaseDatos.mdb")
End Sub 
 */
	}

	 void examinarArchivo ( String nomArchivo ) {

/* 
 

Sub examinarArchivo(nomArchivo As String)
    On Error GoTo ErrorProducido
    Dim fso As Object 'New FileSystemObject
    Dim archivo As Object 'File
    Dim mensaje As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivo = fso.GetFile(nomArchivo)
    
    'mensaje = vbCrLf & "Fecha creación: " _
        & archivo.DateCreated _
        & vbCrLf & "Fecha último acceso: " _
        & archivo.DateLastAccessed _
        & vbCrLf & "Fecha última modificación: " _
        & archivo.DateLastModified _
        & vbCrLf & "Letra unidad: " _
        & archivo.Drive _
        & vbCrLf & "Nombre archivo: " _
        & archivo.Name _
        & vbCrLf & "Carpeta: " _
        & archivo.ParentFolder _
        & vbCrLf & "Ruta completa: " _
        & archivo.Path _
        & vbCrLf & "Nombre archivo DOS: " _
        & archivo.ShortName _
        & vbCrLf & "Nombre ruta DOS: " _
        & archivo.ShortPath _
        & vbCrLf & "Tamaño: " _
        & (Format(archivo.Size, "##,###")) / 1024 & " KB" _
        & vbCrLf & "Tipo de archivo: " _
        & archivo.Type
        
    Dim tam As Double
    tam = (Format(archivo.Size, "##,###")) / 1024
    tam = tam / 1024
    
    
    mensaje = vbCrLf & "Fecha creación:" & vbTab & vbTab & "" _
        & archivo.DateCreated _
        & vbCrLf & "Fecha último acceso:" & vbTab & "" _
        & archivo.DateLastAccessed _
        & vbCrLf & "Fecha última modificación:" & vbTab & "" _
        & archivo.DateLastModified _
        & vbCrLf & vbCrLf & "Ruta completa:  " _
        & archivo.Path _
        & vbCrLf & vbCrLf & "Tamaño:  " _
        & tam & "  MB" _
        & vbCrLf & vbCrLf & "Tipo de archivo:  " _
        & archivo.Type
        
    Set archivo = Nothing
    Set fso = Nothing
    
    a = MsgBox(vbCrLf & "INFORMACIÓN:" & vbCrLf & mensaje, vbInformation, "INFORMACIÓN DE LA BASE DE DATOS")
Salida:
    Exit Sub
ErrorProducido:
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al recopilar Información")
    Resume Salida
End Sub 
 */
	}

	 void restablecer_Click ( ) {

/* 
 
Private Sub restablecer_Click()
    a = MsgBox("Se restablecerán los datos", vbOKOnly, "Restablecer Datos")
    restab
End Sub 
 */
	}

	 void restab ( ) {

/* 
 


Private Sub restab()
    calendar.Year = Year(Date)
    calendar.Month = Month(Date)
    calendar.Day = Day(Date)
    calendar.Value = fechaAux.Text
    
    AFOROS.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\BaseDatos.mdb;Persist Security Info=False"
    AFOROS.RecordSource = "ESTDATAUX"
    Dim cnn As ADODB.Connection
    On Error GoTo err_restablecer
    Set cnn = New ADODB.Connection
        With cnn
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .ConnectionString = "Data Source=" & App.Path & "\BaseDatos.mdb"
            .Open
            .Execute "DELETE * FROM ESTDATAUX;"
            SQL = "INSERT INTO ESTDATAUX " & _
            " (IDEST, FECHAMED, HORAMED, TOTVEH1, VELMED1, VEHPES1, TOTVEH2, VELMED2, VEHPES2, TOTVEHT, VELMEDT, VEHPEST) " & _
            " SELECT A.IDEST, A.FECHAMED, A.HORAMED, A.TOTVEH, A.VELMED, A.VEHPES, B.TOTVEH, B.VELMED, B.VEHPES, A.TOTVEH+B.TOTVEH, ((A.VELMED+B.VELMED)/2), A.VEHPES+B.VEHPES " & _
            " FROM ESTDAT AS A, ESTDAT AS B " & _
            " WHERE A.IDEST= B.IDEST AND A.FECHAMED=B.FECHAMED AND A.HORAMED=B.HORAMED AND A.CANAL=1 AND B.CANAL=2;"
            .Execute SQL
            .Close
        End With
    AFOROS.Refresh
    
    progreso.Value = 0
    
    estText.Text = ""
    loc.Visible = False
    locText.Visible = False
    pk.Visible = False
    pkText.Visible = False
    tipol.Visible = False
    tipoText.Visible = False
    del.Visible = False
    delText.Visible = False
    tra.Visible = False
    traText.Visible = False
    c1.Visible = False
    c1Text.Visible = False
    c2.Visible = False
    c2Text.Visible = False
    c3.Visible = False
    c3Text.Visible = False
    c4.Visible = False
    c4Text.Visible = False
    
salir:
    Exit Sub
err_restablecer:
    a = MsgBox("ERROR: " & Err.Description, vbExclamation, "Error al Restablecer")
    cnn.Close
    Resume salir
End Sub 
 */
	}

	 void salir_Click ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void fechaAux ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void GetSetting ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void MsgBox ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void App ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void Left ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void archivo ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void Mid ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void Len ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void cont ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void ini ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void hora ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void acumulado ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void ano ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void mes ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void dia ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void Year ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void Month ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void Day ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void calendar ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

	 void tam ( ) {

/* 
 Private Sub salir_Click()
    a = MsgBox("¿Desea salir de la aplicación?", vbOKCancel, "Salir")
    If a = 1 Then
        Unload Forms(this)
    End If
End Sub 
 */
	}

}
