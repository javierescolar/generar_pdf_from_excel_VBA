Private Sub Workbook_Open()
    'constantes'
    Dim HOJA_EXP, HOJA_EXP_DET, PLANTILLA_RESUMIDA, PLANTILLA_ALBARAN, CONFIGURABLES, RUTA_FICHEROS_TMP As String
    HOJA_EXP = "expediciones"
    HOJA_EXP_DET = "expediciones_lineas"
    PLANTILLA_RESUMIDA = "plantilla_resumida"
    PLANTILLA_ALBARAN = "plantilla_albaran"
    CONFIGURABLES = "configurables"
    RUTA_FICHEROS = "C:\Users\ES3756\Documents\EXPEDICIONES\PDF\"
    'variables necesarias par ala logica del programa'
    Dim uniques As Collection
    Dim servicios_expedicion As Range
    Dim remplazar As Object, tmp
    Dim posicion_principal, posicion_final, posicion_principal_serv_linea, posicion_final_serv_linea As Integer
    Dim fila, celda As Integer
    Dim id_servicio_linea, num_albaran As Long
    'EMPEZAMOS LA GENERACION DE LA LOGICA'
    Set uniques = GetUniqueValues(Worksheets(HOJA_EXP).Range("A:A").Value)
    For Each id_tecnico In uniques
        'RESET DE APP'
        resetApp HOJA_EXP, HOJA_EXP_DET, PLANTILLA_RESUMIDA, PLANTILLA_ALBARAN, CONFIGURABLES
        If id_tecnico <> "id_tecnico" Then
             Set remplazar = CreateObject("Scripting.Dictionary")
             'hacemos una copia de la plantilla'
             copiarHoja PLANTILLA_RESUMIDA, id_tecnico & "exp", HOJA_EXP
             posicion_principal = calcularPrimeraPosicion(HOJA_EXP, "A:A", id_tecnico)
             posicion_final = calcularUltimaPosicion(HOJA_EXP, "A:A", id_tecnico)
             'Anadimos las variables que han de sustituirse en elbaran'
             remplazar.Add "[%id_tecnico%]", Worksheets(HOJA_EXP).Cells(posicion_principal, 1)
             remplazar.Add "[%nom_tecnico%]", Worksheets(HOJA_EXP).Cells(posicion_principal, 2)
             remplazar.Add "[%id_expedicion%]", Worksheets(HOJA_EXP).Cells(posicion_principal, 3)
             'remplazamos las variables del albaran creado'
             For Each r In remplazar
                remplazarCadenas id_tecnico & "exp", r, remplazar(r)
             Next r
             'obtenemos el rango de servicios en la expedicion'
             Set servicios_expedicion = Worksheets(HOJA_EXP).Range("D" & posicion_principal, "P" & posicion_final)
             'añadimos las filas al albaran creado LE SUMAMOS LAS FILAS NECESARIAS PARA LA PLANTILLA'
             'recorremos la matriz de las filas'
             fila = 10
             For i = 1 To servicios_expedicion.Rows.Count
                celda = 2
                'recorremos la matriz de las columnas'
                For j = 1 To servicios_expedicion.Columns.Count
                   'ultimo campo(observaciones) debajo del servicio para ganar espacio'
                   If servicios_expedicion.Columns.Count = j Then
                       fila = fila + 1
                       Worksheets(id_tecnico & "exp").Cells(fila, 2) = "Obs:"
                       Worksheets(id_tecnico & "exp").Cells(fila, 3) = servicios_expedicion.Cells(i, j)
                       Worksheets(id_tecnico & "exp").Range(Cells(fila, 3), Cells(fila, servicios_expedicion.Columns.Count)).Merge
                       With Worksheets(id_tecnico & "exp").Range(Cells(fila, 2), Cells(fila, servicios_expedicion.Columns.Count)).Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = 1
                        End With
                   Else
                       Worksheets(id_tecnico & "exp").Cells(fila, celda) = servicios_expedicion.Cells(i, j)
                   End If
                   celda = celda + 1
                Next j
                fila = fila + 1
             Next i
             'Crear PDF'
             exportarPdf id_tecnico & "exp", RUTA_FICHEROS
             'encriptamos el pdf con zip'
             'encryptFile RUTA_FICHEROS, id_tecnico & "exp"
             'Enviar Mail'
             'enviarMail Worksheets(CONFIGURABLES).Cells(calcularPrimeraPosicion(CONFIGURABLES, "A:A", id_tecnico), 2), "Expedicion Medio", "Saludos.", RUTA_FICHEROS & id_tecnico & "exp" & ".zip"
            'comprobamos si el tecnico tiene asignada la generacion de albaranes'
            If Worksheets(CONFIGURABLES).Cells(calcularPrimeraPosicion(CONFIGURABLES, "A:A", id_tecnico), 4) = 1 Then
                 posicion_principal = calcularPrimeraPosicion(HOJA_EXP_DET, "A:A", id_tecnico)
                 posicion_final = calcularUltimaPosicion(HOJA_EXP_DET, "A:A", id_tecnico)
                 For pos_servicio = posicion_principal To posicion_final
                    If id_servicio_linea <> Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 3) Then
                        'setteamos el servicio para calcular las lineas del albaran'
                         id_servicio_linea = Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 3)
                         'hacemos una copia de la plantilla'
                         num_albaran = Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 4)
                        copiarHoja PLANTILLA_ALBARAN, id_tecnico & "-albaran-" & num_albaran, HOJA_EXP_DET
                         'Anadimos las variables que han de sustituirse en elbaran'
                        Set remplazar = CreateObject("Scripting.Dictionary")
                        remplazar.Add "[%telefono_servicio%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 6)
                        remplazar.Add "[%paciente%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 7)
                        remplazar.Add "[%direccion%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 8)
                        remplazar.Add "[%cod_postal%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 9)
                        remplazar.Add "[%poblacion%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 10)
                        remplazar.Add "[%telefono%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 11)
                        remplazar.Add "[%num_afiliacion%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 12)
                        remplazar.Add "[%id_expedicion%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 2)
                        remplazar.Add "[%num_albaran%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 4)
                        remplazar.Add "[%orden%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 5)
                        remplazar.Add "[%id_tratamiento%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 13)
                        remplazar.Add "[%terapia%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 14)
                        remplazar.Add "[%comentarios%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 15)
                        remplazar.Add "[%tipo_servicio%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 16)
                        remplazar.Add "[%tipo_mascara%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 17)
                        remplazar.Add "[%observaciones%]", Worksheets(HOJA_EXP_DET).Cells(pos_servicio, 18)
                        'remplazamos las variables del albaran creado'
                        For Each r In remplazar
                           remplazarCadenas id_tecnico & "-albaran-" & num_albaran, r, remplazar(r)
                        Next r
                        'añadimos las filas al albaran creado LE SUMAMOS LAS FILAS NECESARIAS PARA LA PLANTILLA'
                        fila = 17
                        celda = 2
                        'obtenemos el rango de servicios en la expedicion'
                        Set servicios_expedicion = Worksheets(HOJA_EXP_DET).Range("S" & pos_servicio, "AB" & calcularUltimaPosicion(HOJA_EXP_DET, "C:C", id_servicio_linea))
                        'recorremos la matriz de las filas'
                        For i = 1 To servicios_expedicion.Rows.Count
                            celda = 2
                           'recorremos la matriz de las columnas'
                           For j = 1 To servicios_expedicion.Columns.Count
                              'ultimo campo(observaciones) debajo del servicio para ganar espacio'
                              Worksheets(id_tecnico & "-albaran-" & num_albaran).Cells(fila, celda) = servicios_expedicion.Cells(i, j)
                              celda = celda + 1
                           Next j
                           fila = fila + 1
                        Next i
                        'Crear PDF'
                        exportarPdf id_tecnico & "-albaran-" & num_albaran, RUTA_FICHEROS
                    End If
                 Next pos_servicio
            End If
        End If
    Next
End Sub
Sub copiarHoja(hoja_a_copiar, nueva_hoja, HOJA_EXP)
    'If Worksheets(nombreHoja).Name = "" Then
    '   Application.DisplayAlerts = False
    '    Worksheets(nueva_hoja).Delete
    'End If
    Worksheets(hoja_a_copiar).Copy After:=Worksheets(HOJA_EXP)
    ActiveSheet.Name = nueva_hoja
End Sub

Function crearHoja(nombre_hoja)
    Dim wsTest As Worksheet
    Set wsTest = Nothing
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets(nombre_hoja)
    On Error GoTo 0
     
    If wsTest Is Nothing Then
        Worksheets.Add.Name = nombre_hoja
    End If
End Function

Function calcularPrimeraPosicion(nombre_hoja, fila_a_buscar, palabra_buscar) As Integer
    Dim primera_ocurrencia As Integer
    With Worksheets(nombre_hoja).Range(fila_a_buscar)
        primera_ocurrencia = .Find(What:=palabra_buscar, SearchDirection:=xlNext, LookIn:=xlValues).Row
    End With
    calcularPrimeraPosicion = primera_ocurrencia
End Function

Function calcularUltimaPosicion(nombre_hoja, fila_a_buscar, palabra_buscar) As Integer
    Dim ultima_ocurrencia As Integer
    With Worksheets(nombre_hoja).Range(fila_a_buscar)
       ultima_ocurrencia = .Find(What:=palabra_buscar, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
    End With
    calcularUltimaPosicion = ultima_ocurrencia
End Function

Sub remplazarCadenas(nombre_hoja, variable, texto)
    Worksheets(nombre_hoja).Cells.Replace What:=variable, Replacement:=texto
End Sub

Public Function GetUniqueValues(ByVal values As Variant) As Collection
    Dim result As Collection
    Dim cellValue As Variant

    Set result = New Collection
    Set GetUniqueValues = result

    On Error Resume Next

    For Each cellValue In values
        If Trim(cellValue) = "" Then GoTo NextValue
        result.Add Trim(cellValue), Trim(cellValue)
NextValue:
        Next cellValue

    On Error GoTo 0
    
End Function

Sub resetApp(HOJA_EXP, HOJA_EXP_DET, PLANTILLA, PLANTILLA_ALBARAN, CONFIGURABLES)
    For Each hoja In Worksheets
    If hoja.Name <> HOJA_EXP And hoja.Name <> PLANTILLA And hoja.Name <> PLANTILLA_ALBARAN And hoja.Name <> CONFIGURABLES And hoja.Name <> HOJA_EXP_DET Then
        Application.DisplayAlerts = False
        hoja.Delete
    End If
    Next hoja
End Sub

Sub enviarMail(destinatario, asunto, mensaje, ruta_fichero)
    Dim oApp As Object
    Dim oMail As Object
    'Create and show the Outlook mail item
    Set oApp = CreateObject("Outlook.Application")
    Set oMail = oApp.CreateItem(0)
    With oMail
        .To = destinatario
        .Subject = asunto
        .body = mensaje
        .Attachments.Add ruta_fichero
        .Send
    End With
End Sub

Sub exportarPdf(hoja, ruta_pdf)
    Worksheets(hoja).ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=ruta_pdf & "\" & hoja & ".pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub


Sub encryptFile(ruta_dir, archivo_hoja)
    strDestFileName = ruta_dir & archivo_hoja & ".zip"
    strSourceFileName = ruta_dir & archivo_hoja & ".pdf"
    str7ZipPath = "C:\Program Files\7-Zip\7z.exe"
    strPassword = "linde"
    strCommand = str7ZipPath & " -p" & strPassword & " a -tzip """ & strDestFileName & """ """ & strSourceFileName & """"
    Shell strCommand
End Sub







