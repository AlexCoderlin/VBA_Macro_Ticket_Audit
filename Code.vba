'-------------------------------
'   Macro Ticket Audit
'-------------------------------
Sub Ticket_Audit()

'Deshabilitamos la actualizacion grafica para reducir recursos durante la ejecucion
Application.ScreenUpdating = False
Application.DisplayAlerts = False

    'Inicializacion de valiables Locales.
    Dim WB_Macro_TAP, WB_Export As Workbook
    Dim Header_Array(), Agents_Array(), NameSearch, UltCol, UltColLet, CORRECTO, ERROR As Variant
    Dim Number_of_Items, Number_of_Items2, Number_in_Array, Number_in_Array2, TotalRows As Integer
    Dim y As Object

    'Almacenamos el libro actual en variable
    Set WB_Macro_TAP = Application.ThisWorkbook

    'Si tabla de hoja Data tiene datos, borramos datos. Si esta vacia, la dejamos asi.
        If Not Sheets("Data").ListObjects("TableData").DataBodyRange Is Nothing Then

            Sheets("Data").ListObjects("TableData").DataBodyRange.Delete
            Sheets("Data").ListObjects("TableData").ListRows.Add AlwaysInsert:=False

        End If

    'Si tabla de hoja Inicio tiene datos, borramos datos. Si esta vacia, la dejamos asi.
        If Not Sheets("Inicio").ListObjects("Table_Inicio").DataBodyRange Is Nothing Then

            Sheets("Inicio").ListObjects("Table_Inicio").DataBodyRange.Delete
            'Sheets(1).ListObjects("Table_Inicio").ListRows.Add AlwaysInsert:=True

        End If   

'---- Llenamos Header_Array con nombre de columnas indispensables para el export. ----

    'Asignamos el tamaño de la lista que contiene las columnas necesarias para el export (Descrita en hoja "Info").
    Number_of_Items = Sheets("Info").Range("G10").End(xlDown).Row - 9

    'Redimencionamos arreglo con el tamaño de la lista.
    Number_in_Array = Number_of_Items -1 'Menos 1 por el 0 del arreglo
    ReDim Header_Array(Number_in_Array)

    'Preparamos arreglo con valores de la lista con columnas necesarias para el export.
        For i = 0 To Number_in_Array

            Header_Array(i) = Sheets("Info").Range("G" & 10 + i).Value

        Next i

'---- Llenamos Agents_Array con nombres de agentes. ----

    'Asignamos el tamaño de la lista que contiene las columnas necesarias para el export (Descrita en hoja "Info").
    Number_of_Items2 = Sheets("Agentes").Range("A2").End(xlDown).Row - 1

    'Redimencionamos arreglo con el tamaño de la lista.
    Number_in_Array2 = Number_of_Items2 - 1 'Menos 1 por el 0 tiene el arreglo
    ReDim Agents_Array(Number_in_Array2)

    'Preparamos arreglo con valores de la lista con columnas necesarias para el export.
        For i = 0 To Number_in_Array2

            Agents_Array(i) = Sheets(2).Range("A" & 2 + i).Value

        Next i

'---- Abrimos libro con Export y trabajamos en el, importamos resultados y cerramos libro ----

    MsgBox "Por favor seleccione: EXPORT", vbExclamation, "Abrir Datos"

    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "Export", "*.xlsx; *.xlsm; *.csv"
        .AllowMultiSelect = False
        .Show
            If .SelectedItems.Count > 0 Then
                Application.Workbooks.Open .SelectedItems(1)
                Set WB_Export = Application.ActiveWorkbook

                'Calculamos el tamaño de datos en export.
                TotalRows = Range("A2").End(xlDown).Row 'Encuentra ultima fila.
                UltCol = Range("A1").End(xlToRight).Address
                UltColLet = Mid(UltCol, InStr(UltCol, "$") + 1, InStr(2, UltCol, "$") - 2) 'Encuentra ultima columna.

        '-------- Buscamos si las columnas del export contienen datos del arreglo Header_Array.
                For i = LBound(Header_Array) To UBound(Header_Array)
                                        
                    NameSearch = Header_Array(i)
                    Set y = Range("A1:" & UltColLet & "1").Find(What:=NameSearch)
                                                
                    If y is Nothing Then

                        MsgBox ("Este Export NO contiene las columnas necesarias para esta macro. Falta valor: " & NameSearch), vbExclamation, "ERROR"
                        WB_Export.Close SaveChanges:=False
                        End                          

                    End If
                            
                Next i

        '-------- Reducimos datos a importar filtrando solo los realizados por los agentes.
                'Creamos tabla en libro export
                WB_Export.Sheets(1).ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = "TablaExport"

                'Filtramos tabla con los agentes. 'Columna correspondiente a Agentes (Open By) es 6.
                WB_Export.Sheets(1).ListObjects("TablaExport").Range.AutoFilter _
                Field:=6, _ 
                Criteria1:=Agents_Array(), _
                Operator:=xlFilterValues

        '-------- Copiamos solo los datos que necesitamos.
                For i = LBound(Header_Array) To UBound(Header_Array)
                                        
                    NameSearch = Header_Array(i)

                    IF NameSearch = "Call #" Then

                        NameSearch = "Call '#"   'Se agrega "'" en caracteres especiales sino genera error.                        

                    End If

                    WB_Export.Sheets(1).Range("TablaExport[" & NameSearch & "]").SpecialCells(xlCellTypeVisible).Copy
                    WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("" & Header_Array(i)).DataBodyRange.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                        ' With WB_Macro_TAP.Sheets("Data").ListObjects("TableData")
                                'If Not .DataBodyRange Is Nothing Then .ListColumns("Opened By").DataBodyRange.PasteSpecial xlPasteValuesAndNumberFormats
                        'End With

                Next i

                Application.CutCopyMode = False        'Libera portapapeles
                WB_Export.Close SaveChanges:=False    'Cierra Libro Export

            Else

                MsgBox "Usted no selecciono correctamente un Export.", vbExclamation, "ERROR"
                End

            End If
    End With

'---- Trabajamos en calculos de Tabla en hoja Data. ----

    'Podemos cambiar a cualquier necesidad las variables siguientes para formular en celdas:
    CORRECTO = 0
    ERROR = 1

    'Concatena columnas F, G y H para validar mismo patron con hoja "Templates".
     WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("CSP_Assistant").DataBodyRange.Formula = "=SUBSTITUTE([@Category]&[@Subcategory]&[@[Product Type]],"" "","""")"

    'Concatena columnas F, H e I para validar mismo patron con hoja "Templates".
     WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("CSI_Assistant").DataBodyRange.Formula = "=SUBSTITUTE([@Category]&[@[Product Type]]&[@[Issue Type]],"" "","""")"

    'Extrae y concatena solo los digitos de columna "RB Phone", valida si son 8 o mas digitos para numero telefonico correcto. (solo si "Call source" = PHONE).
     WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("Val_Phone").DataBodyRange.Formula2R1C1 = "=IF([@[Call Source]]=""PHONE"",IF(LEN(CONCAT(IF(ISNUMBER(MID([@[RB Phone]],ROW(INDIRECT(""1:""&LEN([@[RB Phone]]))),1)*1),MID([@[RB Phone]],ROW(INDIRECT(""1:""&LEN([@[RB Phone]]))),1),"""")))>=8," & CORRECTO &  "," & ERROR & "),0)"

    'Si "Quick Call Id" vacio entonces buscar si hay un Quick Id disponible en lista de templates, si hay disponible califica.
     WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("Val_Quick").DataBodyRange.Formula = "=IF([@[Quick Call ID]]="""",IF(IFNA(VLOOKUP([@[CSP_Assistant]],Table7[[CSP_Assistant]:[unique.id2]],2,FALSE),13)=13," & CORRECTO & "," & ERROR &"),0)"
     'WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("Val_Quick").DataBodyRange.Formula = "=IF([@[Quick Call ID]]="""",IF(IFNA(VLOOKUP([@[CSP_Assistant]],Table7[[CSP_Assistant]:[unique.id2]],2,FALSE),13)<>13,""FALTA Quick"",""Quick no necesario""),0)"

    'Columna "Knowledgebase ID" debe contener >= 13 digitos.
     WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("Val_KB").DataBodyRange.Formula = "=IF(LEN([@[Knowledgebase ID]])>=13," & CORRECTO &  "," & ERROR & ")"

    'Suma todos los errores detectados en el ticket.
     WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("Error_Total").DataBodyRange.Formula = "=SUM(TableData[@[Val_Phone]:[Val_KB]])"

    'Si ticket tiene 1 o mas errores se contabiliza como ticket con error.
     WB_Macro_TAP.Sheets("Data").ListObjects("TableData").ListColumns("Error_Ticket").DataBodyRange.Formula = "=IF([@[Error_Total]]>0,1,0)"

    'Actualizamos tablas dinamicas con cambios.
     ActiveWorkbook.RefreshAll

'---- Trabajamos en resultados de hoja Inicio. ----

    'Sincronizamos agentes a tabla de hoja inicio.
    WB_Macro_TAP.Sheets("Agentes").ListObjects("TableAgents").ListColumns(1).DataBodyRange.Copy
    WB_Macro_TAP.Sheets("Inicio").Range("D18").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    'Reutilizamos variable para calcular tamaño de Tabla dinamica.
     TotalRows = ActiveWorkbook.Sheets(4).Range("B5").End(xlDown).Row - 1

    'Formulamos Columna 2 de Table_Inicio.
     WB_Macro_TAP.Sheets("Inicio").ListObjects("Table_Inicio").ListColumns(2).DataBodyRange.Formula = "=IFERROR(GETPIVOTDATA(""Count of Call Source"",Dinamicas!$B$4, ""Opened by"",[@[Nombre:]]),0)"

    'Formulamos Columna 3 de Table_Inicio.
     WB_Macro_TAP.Sheets("Inicio").ListObjects("Table_Inicio").ListColumns(3).DataBodyRange.Formula = "=IFERROR(GETPIVOTDATA(""Sum of Error_Ticket"",Dinamicas!$B$4, ""Opened by"",[@[Nombre:]]),0)"
  
    'Asignamos formato porcentaje.
    'ActiveWorkbook.Sheets("Inicio").ListObjects("Table_Inicio").ListColumns(5).DataBodyRange.NumberFormat = "0.00%"

    'Formulamos Columna 4 de Table_Inicio.
     WB_Macro_TAP.Sheets("Inicio").ListObjects("Table_Inicio").ListColumns(4).DataBodyRange.Formula = "=IFERROR(GETPIVOTDATA(""Sum of Error_Total"",Dinamicas!$B$4, ""Opened by"",[@[Nombre:]]),0)"

    'Formulamos Columna 5 de Table_Inicio.
     WB_Macro_TAP.Sheets("Inicio").ListObjects("Table_Inicio").ListColumns(5).DataBodyRange.Formula = "=IFERROR(([@[Tickets generados]]-[@[Tickets con Error]])/[@[Tickets generados]],0)"
    
    'Formulamos Columna 5 de Table_Inicio.     
     WB_Macro_TAP.Sheets("Inicio").ListObjects("Table_Inicio").ListColumns(6).DataBodyRange.Formula = "=ROUND([@Efectividad]*100,0)"
     WB_Macro_TAP.Sheets("Inicio").ListObjects("Table_Inicio").ListColumns(6).DataBodyRange.NumberFormat = "0"

    'Actualizamos tablas dinamicas con cambios.
     ActiveWorkbook.RefreshAll
     WB_Macro_TAP.Sheets("Inicio").Activate
     Range("M5").Select

     MsgBox "Export finalizado con exito.",, "LISTO"

    'Encendemos la actualizacion grafica.
     Application.DisplayAlerts = True
     Application.ScreenUpdating = True

End Sub
