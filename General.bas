Attribute VB_Name = "General"
Sub conGranAmorParaTi()

    Dim tiempoInicio As Double
    Dim tiempoFinal As Double
    Dim duracion As Double

    Dim fechaInicial As String, fechaCierre As String
    Dim mesInicial, diaInicial, yearInicial, mesFinal, diaFinal, yearFinal
    Dim mesInicialPalabrbas, mesFinalPalabras
    Dim observaciones As String
    Dim i As Integer
    Dim valorCelda, vvalorCeldaDos As String
    Dim valorCeldaCantidad As Integer, valorCeldaCantidadDos As Integer, valorSumatoria As Integer
    Dim valorSocio2, valorSocio3, valorSocio2_2, valorSocio3_3 As Integer
    Dim valorSumatoriaUno As Integer, valorSumatoriaDos As Integer
    Dim tablaDnamica As PivotTable
    Dim sumaChaquetaImper As Integer, sumaPantalonImper As Integer
    
    ' Registra el tiempo de inicio
    tiempoInicio = Timer
    
    fechaInicial = ThisWorkbook.Sheets("PREFACTURA").Range("N4").Value
    fechaCierre = ThisWorkbook.Sheets("PREFACTURA").Range("N5").Value
    
    diaInicial = Split(fechaInicial, "/")
    diaInicial = diaInicial(0)
    mesInicial = Split(fechaInicial, "/")
    mesInicial = mesInicial(1)
    yearInicial = Split(fechaInicial, "/")
    yearInicial = yearInicial(2)
    
    diaFinal = Split(fechaCierre, "/")
    diaFinal = diaFinal(0)
    mesFinal = Split(fechaCierre, "/")
    mesFinal = mesFinal(1)
    yearFinal = Split(fechaCierre, "/")
    yearFinal = yearFinal(2)
    
    Set tablaDnamica = ThisWorkbook.Sheets("RESUMEN").PivotTables("LuisAmortegui")
    ' Actualiza la tabla dinámica
    tablaDnamica.RefreshTable
    
    ThisWorkbook.Sheets("PREFACTURA").Range("B16").Value = mesInicial & "/" & diaInicial & "/" & yearInicial
    ThisWorkbook.Sheets("PREFACTURA").Range("E16").Value = mesFinal & "/" & diaFinal & "/" & yearFinal
    
    Select Case mesInicial
    Case "01"
        mesInicialPalabras = "Enero"
    Case "02"
        mesInicialPalabras = "Febrero"
    Case "03"
        mesInicialPalabras = "Marzo"
    Case "04"
        mesInicialPalabras = "Abril"
    Case "05"
        mesInicialPalabras = "Mayo"
    Case "06"
        mesInicialPalabras = "Junio"
    Case "07"
        mesInicialPalabras = "Julio"
    Case "08"
        mesInicialPalabras = "Agosto"
    Case "09"
        mesInicialPalabras = "Septiembre"
    Case "10"
        mesInicialPalabras = "Octubre"
    Case "11"
        mesInicialPalabras = "Noviembre"
    Case "12"
        mesInicialPalabras = "Diciembre"
    Case Else
        ' Lógica por defecto si mesInicial no coincide con ninguno de los casos anteriores
    mesInicialPalabras = "Mes no válido"
    End Select
    
    
    Select Case mesFinal
    Case "01"
        mesFinalPalabras = "Enero"
    Case "02"
        mesFinalPalabras = "Febrero"
    Case "03"
        mesFinalPalabras = "Marzo"
    Case "04"
        mesFinalPalabras = "Abril"
    Case "05"
        mesFinalPalabras = "Mayo"
    Case "06"
        mesFinalPalabras = "Junio"
    Case "07"
        mesFinalPalabras = "Julio"
    Case "08"
        mesFinalPalabras = "Agosto"
    Case "09"
        mesFinalPalabras = "Septiembre"
    Case "10"
        mesFinalPalabras = "Octubre"
    Case "11"""
        mesFinalPalabras = "Noviembre"
    Case "12"
        mesFinalPalabras = "Diciembre"
    Case Else
        ' Lógica por defecto si mesInicial no coincide con ninguno de los casos anteriores
    mesFinalPalabras = "Mes no válido"
    End Select
    
    observaciones = "OBSERVACIONES: Lavado de prendas del " & diaInicial & " de " & mesInicialPalabras & " al " & diaFinal & " de " & mesFinalPalabras & " del " & yearFinal
    
    ThisWorkbook.Sheets("PREFACTURA").Range("B23").Value = observaciones
    
    ' Validacion Pantalon Jena mas pantalon termico
    For i = 1 To 20
        valorCelda = ThisWorkbook.Sheets("RESUMEN").Range("A" & i).Value
        
        If valorCelda = "Pantalon Jean" Then
            valorCeldaCantidad = ThisWorkbook.Sheets("RESUMEN").Range("B" & i).Value
            valorSocio2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & i).Value
            valorSocio3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & i).Value
            For j = 1 To 20
                vvalorCeldaDos = ThisWorkbook.Sheets("RESUMEN").Range("A" & j).Value
                If vvalorCeldaDos = "Pantalon Termico" Then
                    valorCeldaCantidadDos = ThisWorkbook.Sheets("RESUMEN").Range("B" & j).Value
                    valorSocio2_2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & j).Value
                    valorSocio3_3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & j).Value
                End If
            Next j
            
            valorSumatoria = valorCeldaCantidad + valorCeldaCantidadDos
            valorSumatoriaUno = valorSocio2 + valorSocio2_2
            valorSumatoriaDos = valorSocio3 + valorSocio3_3
            ThisWorkbook.Sheets("PREFACTURA").Range("E27").Value = valorSumatoria
            ThisWorkbook.Sheets("PREFACTURA").Range("F27").Value = valorSumatoriaUno
            ThisWorkbook.Sheets("PREFACTURA").Range("G27").Value = valorSumatoriaDos
        End If
    Next i
    
    i = 0
    valorCeldaCantidad = 0
    valorSocio2 = 0
    valorSocio3 = 0
    
    ' Validacion camisa polo
    For i = 1 To 20
        valorCelda = ThisWorkbook.Sheets("RESUMEN").Range("A" & i).Value
        
        If valorCelda = "Camisa Polo" Then
            valorCeldaCantidad = ThisWorkbook.Sheets("RESUMEN").Range("B" & i).Value
            valorSocio2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & i).Value
            valorSocio3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & i).Value
            ThisWorkbook.Sheets("PREFACTURA").Range("E28").Value = valorCeldaCantidad
            ThisWorkbook.Sheets("PREFACTURA").Range("F28").Value = valorSocio2
            ThisWorkbook.Sheets("PREFACTURA").Range("G28").Value = valorSocio3
        End If
    Next i
    
    ' Validacion Buso
    For i = 1 To 20
        valorCelda = ThisWorkbook.Sheets("RESUMEN").Range("A" & i).Value
        
        If valorCelda = "Buso" Then
            valorCeldaCantidad = ThisWorkbook.Sheets("RESUMEN").Range("B" & i).Value
            valorSocio2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & i).Value
            valorSocio3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & i).Value
            ThisWorkbook.Sheets("PREFACTURA").Range("E29").Value = valorCeldaCantidad
            ThisWorkbook.Sheets("PREFACTURA").Range("F29").Value = valorSocio2
            ThisWorkbook.Sheets("PREFACTURA").Range("G29").Value = valorSocio3
        End If
    Next i
    
    ' Validacion pantalon chaqueta impermeable
    For i = 1 To 20
        valorCelda = ThisWorkbook.Sheets("RESUMEN").Range("A" & i).Value
        If valorCelda = "Chaqueta Impermeable" Then
            valorCeldaCantidad = ThisWorkbook.Sheets("RESUMEN").Range("B" & i).Value
            valorSocio2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & i).Value
            valorSocio3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & i).Value
        End If
        
        For j = 1 To 20
                vvalorCeldaDos = ThisWorkbook.Sheets("RESUMEN").Range("A" & j).Value
                If vvalorCeldaDos = "Pantalon Impermeable" Then
                    valorCeldaCantidadDos = ThisWorkbook.Sheets("RESUMEN").Range("B" & j).Value
                    valorSocio2_2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & j).Value
                    valorSocio3_3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & j).Value
                End If
        Next j
    Next i
    
    sumaChaquetaImper = valorCeldaCantidad + valorSocio2 + valorSocio3
    sumaPantalonImper = valorCeldaCantidadDos + valorSocio2_2 + valorSocio3_3
    
    If sumaChaquetaImper > sumaPantalonImper Then
        ThisWorkbook.Sheets("PREFACTURA").Range("E30").Value = valorCeldaCantidad
        ThisWorkbook.Sheets("PREFACTURA").Range("F30").Value = valorSocio2
        ThisWorkbook.Sheets("PREFACTURA").Range("G30").Value = valorSocio3
    ElseIf sumaPantalonImper > sumaChaquetaImper Then
        ThisWorkbook.Sheets("PREFACTURA").Range("E30").Value = valorCeldaCantidadDos
        ThisWorkbook.Sheets("PREFACTURA").Range("F30").Value = valorSocio2_2
        ThisWorkbook.Sheets("PREFACTURA").Range("G30").Value = valorSocio3_3
    ElseIf sumaChaquetaImper = sumaPantalonImper Then
        ThisWorkbook.Sheets("PREFACTURA").Range("E30").Value = valorCeldaCantidad
        ThisWorkbook.Sheets("PREFACTURA").Range("F30").Value = valorSocio2
        ThisWorkbook.Sheets("PREFACTURA").Range("G30").Value = valorSocio3
    End If
    
    ' Validacion chaqueta
    For i = 1 To 20
        valorCelda = ThisWorkbook.Sheets("RESUMEN").Range("A" & i).Value
        
        If valorCelda = "Chaqueta" Then
            valorCeldaCantidad = ThisWorkbook.Sheets("RESUMEN").Range("B" & i).Value
            valorSocio2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & i).Value
            valorSocio3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & i).Value
            ThisWorkbook.Sheets("PREFACTURA").Range("E31").Value = valorCeldaCantidad
            ThisWorkbook.Sheets("PREFACTURA").Range("F31").Value = valorSocio2
            ThisWorkbook.Sheets("PREFACTURA").Range("G31").Value = valorSocio3
        End If
    Next i
    
    
    ' Validacion Chaleco Reflectivo
    For i = 1 To 20
        valorCelda = ThisWorkbook.Sheets("RESUMEN").Range("A" & i).Value
        
        If valorCelda = "Chaleco Reflectivo" Then
            valorCeldaCantidad = ThisWorkbook.Sheets("RESUMEN").Range("B" & i).Value
            valorSocio2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & i).Value
            valorSocio3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & i).Value
            ThisWorkbook.Sheets("PREFACTURA").Range("E32").Value = valorCeldaCantidad
            ThisWorkbook.Sheets("PREFACTURA").Range("F32").Value = valorSocio2
            ThisWorkbook.Sheets("PREFACTURA").Range("G32").Value = valorSocio3
        End If
    Next i
    
    ' Validacion bata blanca
    For i = 1 To 20
        valorCelda = ThisWorkbook.Sheets("RESUMEN").Range("A" & i).Value
        
        If valorCelda = "Bata" Then
            valorCeldaCantidad = ThisWorkbook.Sheets("RESUMEN").Range("B" & i).Value
            valorSocio2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & i).Value
            valorSocio3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & i).Value
            ThisWorkbook.Sheets("PREFACTURA").Range("E33").Value = valorCeldaCantidad
            ThisWorkbook.Sheets("PREFACTURA").Range("F33").Value = valorSocio2
            ThisWorkbook.Sheets("PREFACTURA").Range("G33").Value = valorSocio3
        End If
    Next i
    
    ' Validacion overol
    For i = 1 To 20
        valorCelda = ThisWorkbook.Sheets("RESUMEN").Range("A" & i).Value
        
        If valorCelda = "Overol" Then
            valorCeldaCantidad = ThisWorkbook.Sheets("RESUMEN").Range("B" & i).Value
            valorSocio2 = ThisWorkbook.Sheets("RESUMEN").Range("C" & i).Value
            valorSocio3 = ThisWorkbook.Sheets("RESUMEN").Range("D" & i).Value
            ThisWorkbook.Sheets("PREFACTURA").Range("E34").Value = valorCeldaCantidad
            ThisWorkbook.Sheets("PREFACTURA").Range("F34").Value = valorSocio2
            ThisWorkbook.Sheets("PREFACTURA").Range("G34").Value = valorSocio3
        End If
    Next i
    
    ' Registra el tiempo final
    tiempoFinal = Timer

    ' Calcula la duración en segundos
    duracion = tiempoFinal - tiempoInicio

    ' Muestra la duración en la ventana de inmediato
    Debug.Print "Tiempo de ejecución: " & duracion & " segundos"

End Sub


