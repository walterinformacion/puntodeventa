VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14880
   OleObjectBlob   =   "UserForm2.frx":0000
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ArchivoPrincipal As Workbook
Dim ArrayAsignaciones As Variant
Dim ArrayFiltroPrincipal As Variant
Public NombreInicio As String
Dim Inicio As Boolean

Public Function RenglonValor(ValorBuscar As Long) As Long
    Dim Renglon As Variant
    Dim RangoPruebas As Range
    Set RangoPruebas = Worksheets("Clientes").Range("A:A")
    On Error GoTo salir
    RenglonValor = Application.WorksheetFunction.Match(ValorBuscar, RangoPruebas, 0)
    Exit Function
    
salir:
    RenglonValor = 0
End Function

Private Sub ComboBoxBusquedaPor_Change()
    
End Sub

Private Sub CommandButton1_Click()
cierreturno.Show
End Sub

Private Sub CommandButton2_Click()
   Inicioturno.Show
End Sub

Private Sub CommandButton20_Click()
   
End Sub


Private Sub CommandButtonAsignacion_Click()

End Sub

Private Sub CommandButton3_Click()
Ventas.Show
End Sub

Private Sub CommandButton5_Click()
End
Unload Me
End Sub



Private Sub CommandButton4_Click()
Dim i As Integer
Dim contador As Integer
Dim porcentaje As Double

With ListBoxAsignaciones
    For i = 0 To .ListCount - 1
    If "Factura" = .List(i, 0) Then
    .ListIndex = i
    
     ListBox1.AddItem
     Me.ListBox1.List(contador, 0) = ListBoxAsignaciones.List(i, 1)
     Me.ListBox1.List(contador, 1) = TextBoxFecha
      Me.ListBox1.List(contador, 2) = TextBoxNombre.Value
      Me.ListBox1.List(contador, 3) = TextBox3.Value
      Me.ListBox1.List(contador, 4) = TextBox4.Value
      Me.ListBox1.List(contador, 5) = TextBoxValor.Value
  
      
      porcentaje = i / (.ListCount - 1)
      Label20.Caption = Format(porcentaje, "0%")
      Me.Label21.Width = (i / (.ListCount - 1)) * Frame1.Width
      
 

   contador = contador + 1
   
   DoEvents
End If
Next
End With

CommandButton6_Click

End Sub

Private Sub CommandButton6_Click()
 li = ListBox1.ListCount - 1
  UltimoRenglon = Worksheets("calculomes").UsedRange.Rows(Worksheets("calculomes").UsedRange.Rows.Count).Row
Rem   Worksheets("calculomes").Range ("A2:j") & UltimoRenglon.Clear
     Sheets("calculomes").Select
    Range("A2").Activate
    n = 2
    
   For i = 0 To li
    ActiveCell.Offset(i, 0) = ListBox1.List(i, 0)
    ActiveCell.Offset(i, 7) = ListBox1.List(i, 1)
    ActiveCell.Offset(i, 2) = ListBox1.List(i, 2)
     ActiveCell.Offset(i, 3) = ListBox1.List(i, 3)
     ActiveCell.Offset(i, 4) = ListBox1.List(i, 4)
     ActiveCell.Offset(i, 5) = ListBox1.List(i, 5)
     ActiveCell.Offset(i, 1) = VBA.Mid(ActiveCell.Offset(i, 7), 18, 7)
     ActiveCell.Offset(i, 6) = "mes:" & VBA.Mid(ActiveCell.Offset(i, 7), 21, 4)
     n = n + 1
    Next
    
      UserForm1.Show
      
End Sub

Private Sub CommandButtonCarpetaAsignaciones_Click()
    ArchivoPrincipal.Activate
    Dim Ruta As String
    Ruta = ThisWorkbook.Path & "\Tools\Ventas"
    ChDrive ("C:")
    ChDir Ruta
    NombredeArchivo = Application.GetOpenFilename("xlsx Files (*.xlsx), *.xlsx")
    Set AplicacionExcel = CreateObject("Excel.Application")
    AplicacionExcel.Visible = True
    If NombredeArchivo <> False Then
        AplicacionExcel.Application.Workbooks.Open NombredeArchivo
    End If
    On Error GoTo salir
     'RutaCompleta = Ruta & "\Tools\Asignaciones\"
     'If Dir(RutaCompleta, vbDirectory) <> "" Then
     '   ActiveWorkbook.FollowHyperlink Address:=RutaCompleta, NewWindow:=True
     'Else
     '   MsgBox "Directorio no existe: " & RutaCompleta
     '   MsgBox "Tus imagenes deben estar en este folder: " & RutaCompleta
     'End If
salir:

End Sub

Private Sub CommandButtonConverttoPDF_Click()
    ArchivoPrincipal.Activate
    If ListBoxAsignaciones.ListIndex >= 0 Then
        ArchivoPrincipal.Activate
        IdReport = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 1)
       
        Worksheets("Factura").ExportAsFixedFormat Type:=xlTypePDF, filename:=NombredeArchivo, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Else
        MsgBox "Seleccione un item.", vbExclamation
    End If
    
'    Dim Ruta As String
'    'Ruta = ThisWorkbook.Path & "\Reportes Varios\Estados de Cuenta"
'    ChDrive ("C:")
'    'ChDir Ruta
'    'NombredeArchivo = Application.GetOpenFilename("pdf Files (*.pdf), *.pdf")
'    NombredeArchivo = ThisWorkbook.Path & "\Tools\Asignaciones con Firma\A" & ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 1) & ".PDF"
'    If Dir(NombredeArchivo) <> "" Then
'        ActiveWorkbook.FollowHyperlink NombredeArchivo
'    End If
End Sub



Private Sub CommandButtonDevolucion_Click()
    ArchivoPrincipal.Activate
    If ListBoxAsignaciones.ListIndex >= 0 Then
        If ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 4) = "VENTA" Then
            Devoluciones.IdReport = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 1)
            Devoluciones.Show
            UserForm_Activate
            If TextBoxTecnico <> "" Then
                TextBoxTecnico_Change
            End If
        End If
    Else
        MsgBox "Seleccione una Venta¡¡¡", vbInformation
    End If
End Sub


Private Sub CommandButtonReporteExcel_Click()
    ArchivoPrincipal.Activate
    If ListBoxAsignaciones.ListIndex >= 0 Then
        ArchivoPrincipal.Activate
        IdReport = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 1)
        ImprimirReporteAsignaciones
    Else
        MsgBox "Seleccione una Venta¡¡¡", vbInformation
    End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Label16_Click()
cierreturno.Show
End Sub

Private Sub Label17_Click()
cierreturno.Show
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub CommandButtonReporteInventario_Click()

End Sub

Private Sub CommandButtonReporteVentas_Click()
    ArchivoPrincipal.Activate
    Reporte.Show
End Sub

Private Sub CommandButtonTicket_Click()
Dim FormatoMonedaCelda As String

FormatoMonedaCelda = "$ ##,##"
  
   
    Dim FormatoTemporal As Double
    Dim FormatoTemporal2 As Double
    Dim FormatoFecha As Date
    Dim FormatoMoneda As Currency
    Dim FormatoMoneda2 As Currency
    Dim cadena As String
    Worksheets("Ticket").Range("A16:Z300").Clear
  
   
       Worksheets("Ticket").Cells(12, 4) = TextBoxAsignacion.Value

       Worksheets("Ticket").Cells(13, 4) = TextBoxFecha.Value
    If TextBoxNombre.Value = "Sin Registro" Then
      Worksheets("Ticket").Cells(11, 2) = "Venta Mostrador"
      Else
       Worksheets("Ticket").Cells(11, 2) = TextBoxNombre.Value
       End If
        Worksheets("Ticket").Cells(12, 2) = TextBox2.Value
        Worksheets("Ticket").Cells(11, 1) = ""
       Worksheets("Ticket").Cells(13, 1) = "NIT:" & TextBox1.Value
      
    For contador = 0 To ListBoxDetalles.ListCount - 1
        Worksheets("Ticket").Cells(contador + 16, 1) = Format(ListBoxDetalles.List(contador, 3), "general number")
        Worksheets("Ticket").Cells(contador + 16, 2) = ListBoxDetalles.List(contador, 1)
        Worksheets("Ticket").Cells(contador + 16, 3) = Format(ListBoxDetalles.List(contador, 4), "$###,###.")
         Worksheets("Ticket").Cells(contador + 16, 3).HorizontalAlignment = xlRight
         
        Worksheets("Ticket").Cells(contador + 16, 4) = ListBoxDetalles.List(contador, 5)
        Worksheets("Ticket").Cells(contador + 16, 4).NumberFormat = FormatoMonedaCelda
        Worksheets("Ticket").Cells(contador + 16, 4).HorizontalAlignment = xlRight
       
        
        
        
        Next
      
        
         Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 16, 1) = "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
          Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 16, 1).Font.Size = 8
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 16, 1).WrapText = True
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 16, 1).VerticalAlignment = xlTop
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 16, 1).HorizontalAlignment = xlCenter
    Range(Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 16, 1), Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 16, 4)).Merge
  
        
       Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 17, 3) = "        Sub-Total"
        Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 17, 3).HorizontalAlignment = xlRight
        FormatoMoneda = Format(TextBox3, "Currency")
        Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 17, 4) = TextBox3
        Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 17, 4).NumberFormat = FormatoMonedaCelda
        
        Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 18, 3) = "        IVA"
        Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 18, 3).HorizontalAlignment = xlRight
        FormatoMoneda = Format(TextBox4, "Currency")
        Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 18, 4) = FormatoMoneda
        Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 18, 4).NumberFormat = FormatoMonedaCelda
    
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 19, 3) = "        Total"
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 19, 3).HorizontalAlignment = xlRight
    
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 19, 4) = TextBoxValor
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 19, 4).NumberFormat = FormatoMonedaCelda
    
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 20, 3) = "        Efectivo"
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 20, 3).HorizontalAlignment = xlRight

    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 20, 4) = Format(TextBoxValor.Value, "Currency")
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 20, 4).NumberFormat = FormatoMonedaCelda
     
     
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 21, 1) = "."
     Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 22, 1) = "."
      Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 23, 1) = "."
   Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 24, 1) = "la siguiente factura en sus efectos a una letra de cambio ART 621,772,773,774 del codigo del comerciante"
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 24, 2).Font.Size = 8
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 24, 2).WrapText = True
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 24, 2).VerticalAlignment = xlTop
    Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 24, 2).HorizontalAlignment = xlCenter
    Range(Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 24, 1), Worksheets("Ticket").Cells(ListBoxDetalles.ListCount + 24, 4)).Merge
  
  Worksheets("Ticket").PrintOut
End Sub

Private Sub CommandButtonWindowConfig_Click()
    ArchivoPrincipal.Activate
    ZConfigWindow.RenglonVentana = 14
    Set ZConfigWindow.VentanaActiva = Me
    ZConfigWindow.Show
End Sub

Private Sub LabelValor_Click()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub ListBoxAsignaciones_Click()
    ArchivoPrincipal.Activate
    'ListBox1.RowSource = ""

    If ListBoxAsignaciones.ListIndex >= 0 Then

            LabelAsignacion = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 0)
            
            If ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 4) = "VENTA" Then
                CommandButtonDevolucion.Visible = True
                CommandButtonConverttoPDF.Visible = True
                CommandButtonReporteExcel.Visible = True
                CommandButtonTicket.Visible = True
                Label6.Visible = True
                Label7.Visible = True
                Label8.Visible = True
                Label13.Visible = True
                
            Else
                CommandButtonDevolucion.Visible = False
                CommandButtonConverttoPDF.Visible = False
                CommandButtonReporteExcel.Visible = False
                CommandButtonTicket.Visible = False
                Label6.Visible = False
                Label7.Visible = False
                Label8.Visible = False
                Label13.Visible = False
            End If
            TextBoxAsignacion = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 1)
            TextBoxNombre = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 2)
            TextBoxKit = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 3)
            TextBoxStatus = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 4)
            TextBoxFecha = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 5)
            TextBoxValor = ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 6)
            'TextBoxEmpresa = Worksheets("DetallesControl").Cells(ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 9), 42)
          
     


    End If
    
    
    ListBoxDetalles.RowSource = ""
    'ImageTool.Picture = Nothing
    Dim ArrayTemporal As Variant
    Dim UltimoRenglon As Long
    If ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 4) = "VENTA" Then
        Worksheets("Temporal2").Activate
        Worksheets("Temporal2").Range("A:Z").Clear
        
        UltimoRenglon = Worksheets("Ventas").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        
        Worksheets("Ventas").Range("A1:Z" & UltimoRenglon).Copy
        Worksheets("Temporal2").Range("A1").PasteSpecial (xlPasteValues)
        
        UltimoRenglon = Worksheets("Temporal2").Range("A" & Rows.Count).End(xlUp).Row
        Worksheets("Temporal2").Range("A1:Z" & UltimoRenglon).AutoFilter FIELD:=1, Criteria1:="<>" & ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 1), VisibleDropDown:=False
        
        UltimoRenglon = Worksheets("Temporal2").Range("A" & Rows.Count).End(xlUp).Row
        Application.DisplayAlerts = False
        Worksheets("Temporal2").Range("A1:Z" & UltimoRenglon).SpecialCells(xlCellTypeVisible).Rows.Delete
        Application.DisplayAlerts = True
        
        UltimoRenglon = Worksheets("Temporal2").Range("A" & Rows.Count).End(xlUp).Row
        
        Worksheets("Temporal2").Range("A:C,F:F,K:Z").Delete
        UltimoRenglon = Worksheets("Temporal2").Range("A" & Rows.Count).End(xlUp).Row
        Set TheRange = Worksheets("Temporal2").Range("A1:J" & UltimoRenglon)
        ArrayTemporal = TheRange
        ListBoxDetalles.List = ArrayTemporal
        
    End If
    
    
    If ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 4) = "DEVOLUCION" Then
        
        Worksheets("Temporal2").Activate
        Worksheets("Temporal2").Range("A:Z").Clear
        
        UltimoRenglon = Worksheets("Devoluciones").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        Worksheets("Devoluciones").Range("A1:Z" & UltimoRenglon).Copy
        Worksheets("Temporal2").Range("A1").PasteSpecial (xlPasteValues)
        UltimoRenglon = Worksheets("Temporal2").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        Worksheets("Temporal2").Range("A1:Z" & UltimoRenglon).AutoFilter FIELD:=1, Criteria1:="<>" & ListBoxAsignaciones.List(ListBoxAsignaciones.ListIndex, 1), VisibleDropDown:=False
        UltimoRenglon = Worksheets("Temporal2").Range("A" & Worksheets("Temporal2").Rows.Count).End(xlUp).Row
        Application.DisplayAlerts = False
        Worksheets("Temporal2").Range("A1:Z" & UltimoRenglon).SpecialCells(xlCellTypeVisible).Rows.Delete
        Application.DisplayAlerts = True
        
        Worksheets("Temporal2").Range("BB1") = "UNO"
        UltimoRenglon = Worksheets("Temporal2").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        
        Worksheets("Temporal2").Range("A:C,F:F,K:Z").Delete
        Worksheets("Temporal2").Range("BB1") = "UNO"
        UltimoRenglon = Worksheets("Temporal2").Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        Set TheRange = Worksheets("Temporal2").Range("A1:J" & UltimoRenglon)
        ArrayTemporal = TheRange
        ListBoxDetalles.List = ArrayTemporal
    End If
    
    If ListBoxDetalles.ListCount > 0 Then
        ListBoxDetalles.Selected(0) = True
    End If
        subtota = 0
        iva = 0
        
        
             For Contadorsuma = 0 To ListBoxDetalles.ListCount - 1
             subtota = subtota + ListBoxDetalles.List(Contadorsuma, 5)
            
             Next
             TextBox3 = Format(subtota, "$ ##,##.##")
             numero = TextBoxValor
              
              numero = RTrim(numero)
         numero = WorksheetFunction.Trim(numero)
         nuemro = LTrim(numero)
         TextBoxValor = numero
         

Dim tamanocadena As Integer
Dim cont As Integer
Dim caracter As String
Dim resultado As String

resultado = ""
tamanocadena = Len(TextBoxValor)

For cont = 1 To tamanocadena - 2
caracter = Mid(TextBoxValor, cont, 1)
If IsNumeric(caracter) Then
resultado = resultado & caracter
End If

Next cont

k = resultado - TextBox3
TextBox4 = Format(k, "currency")

End Sub



Private Sub ListBoxDetalles_Click()
    ArchivoPrincipal.Activate
    Dim ArchivoImagen As String
    ArchivoImagen = ThisWorkbook.Path & "\Tools\Images\" & ListBoxDetalles.List(ListBoxDetalles.ListIndex, 0) & ".JPG"
    If Dir(ArchivoImagen) <> "" Then
        Me.ImageTool.Picture = LoadPicture(ArchivoImagen)
        Me.FrameRefrescar.Repaint
    Else
        Me.ImageTool.Picture = Nothing
        Me.FrameRefrescar.Repaint
    End If
    
End Sub



Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()


End Sub

Private Sub TextBoxAsignacion_Change()

End Sub

Private Sub TextBoxFecha_Change()
TextBoxFecha = Format(TextBoxFecha, "dd/mm/yyyy")
End Sub

Private Sub TextBoxNombre_Change()
  If TextBoxNombre.Value = "Sin Registro" Then
    Me.TextBox1 = Clear
    Me.TextBox2 = Clear
    
  Else
  numero = Sheet5.Range("B" & Rows.Count).End(xlUp).Row
  Me.TextBox1 = Clear
  Me.TextBox2 = Clear

  y = 0
  
  For fila = 2 To numero
     descrip1 = Sheet5.Cells(fila, 2).Value
        If UCase(descrip1) Like "*" & UCase(Me.TextBoxNombre.Value) & "*" Then
       Me.TextBox1 = Sheet5.Cells(fila, 3)
       Me.TextBox2 = Sheet5.Cells(fila, 5)
     
       y = y + 1
       End If
        Next
  End If
  
      
             
               
      
End Sub

Private Sub TextBoxSumaAsignaciones_Change()

End Sub

Private Sub TextBoxTecnico_Change()
    ArchivoPrincipal.Activate
    Dim RangeSuma As Range
    'Application.DisplayAlerts = False
    
    Dim TheRange As Range
    
    ListBoxAsignaciones.Clear
    ListBoxDetalles.Clear
    Set ImageTool.Picture = Nothing
    
    If Trim(TextBoxTecnico) = "" Then
        ListBoxAsignaciones.List = ArrayAsignaciones
        Set RangeSuma = Worksheets("DetallesControl").Range("W:W")
        TextBoxSumaAsignaciones = Format(Application.WorksheetFunction.Sum(RangeSuma), "$ #,##0.00")
        TextBoxConteoAsignaciones = Application.WorksheetFunction.Count(RangeSuma)
        Exit Sub
    End If
    ReDim palabra(100) As String
    Dim ContarPalabra As Long
    palabra(1) = ""
    ContarPalabra = 1
    For ContarLetra = 1 To Len(TextBoxTecnico)
        If Mid(TextBoxTecnico, ContarLetra, 1) = " " Then
            If Not palabra(ContarPalabra) = "" Then
                ContarPalabra = ContarPalabra + 1
                palabra(ContarPalabra) = ""
            End If
        Else
            palabra(ContarPalabra) = palabra(ContarPalabra) & Mid(TextBoxTecnico, ContarLetra, 1)
        End If
    Next
    If palabra(ContarPalabra) = "" Then
        ContarPalabra = ContarPalabra - 1
    End If
    If ContarPalabra > 0 Then
    
        ArchivoPrincipal.Activate
        ThisWorkbook.Application.Visible = False
        
        Worksheets("Temporal").Activate
        Worksheets("Temporal").Range("A:Z").Clear
               
        UltimoRenglon = Worksheets("DetallesControl").Range("AA" & Rows.Count).End(xlUp).Row
        Worksheets("DetallesControl").Range("A2:AZ" & UltimoRenglon).Copy
        Worksheets("Temporal").Range("A2").PasteSpecial (xlPasteValues)
        
        Worksheets("Temporal").Activate
        TotalRenglones = Worksheets("Temporal").Range("A" & Worksheets("Temporal").Rows.Count).End(xlUp).Row
        For Conteo = 1 To ContarPalabra
                Worksheets("Temporal").Range("A1:AZ" & TotalRenglones).AutoFilter FIELD:=8, Criteria1:="<>*" & palabra(Conteo) & "*", VisibleDropDown:=False
                UltimoRenglon = Worksheets("Temporal").Range("A" & Rows.Count).End(xlUp).Row
                If UltimoRenglon > 1 Then
                    Application.DisplayAlerts = False
                    Worksheets("Temporal").Range("A2:AZ" & UltimoRenglon).SpecialCells(xlCellTypeVisible).Rows.Delete
                    Application.DisplayAlerts = True
                End If
        Next
        Cells.AutoFilter
        
        UltimoRenglon = Worksheets("Temporal").Range("A" & Worksheets("Temporal").Rows.Count).End(xlUp).Row
        
        Set TheRange = Worksheets("Temporal").Range("A2:J" & UltimoRenglon)
        ArrayFiltroPrincipal = TheRange
        
        If UltimoRenglon = 1 Then
            ListBoxAsignaciones.Clear
        Else
            ListBoxAsignaciones.List = ArrayFiltroPrincipal
            'ListBoxAsignaciones.Selected(0) = True
        End If
        
    End If

    
    Set RangeSuma = Worksheets("Temporal").Range("W:W")
    TextBoxSumaAsignaciones = Format(Application.WorksheetFunction.Sum(RangeSuma), "$ #,##0.00")
    TextBoxConteoAsignaciones = Application.WorksheetFunction.Count(RangeSuma)
    Set RangeSuma = Worksheets("Temporal").Range("X:X")
    TextBoxSumaDevoluciones = Format(Application.WorksheetFunction.Sum(RangeSuma), "$ #,##0.00")
    
    If ListBoxAsignaciones.ListCount = 0 Then
        
        ListBoxDetalles.RowSource = ""
        LabelAsignacion = Empty
        TextBoxAsignacion = Empty
        TextBoxNombre = Empty
        TextBoxKit = Empty
        TextBoxStatus = Empty
        TextBoxFecha = Empty
        TextBoxValor = Empty
            
    End If
    
End Sub

Private Sub TextBoxTecnico_Enter()
    ArchivoPrincipal.Activate
'    ListBoxDetalles.Clear
'    Me.ImageTool.Picture = Nothing
'    Me.FrameRefrescar.Repaint
    
End Sub







Private Sub TextBoxValor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
End If

End Sub

Private Sub UserForm_Activate()

    
    ArchivoPrincipal.Activate
    If Worksheets("vendedores").Cells(2, 7) = "no" Then
    CommandButton1.Visible = True
    CommandButton3.Visible = True

    Label11.Visible = True
    Label16.Visible = True
    Label17.Visible = True
    Else
    
    CommandButton2.Visible = True
    Label14.Visible = True
    Label15.Visible = True
    
    
    End If
    
   
        
    
    
    Dim PrimerRenglon, UltimoRenglon As Long
    
    Dim RangeBase As Range
    Dim Asignaciones, Devoluciones As Worksheet
    If Not SheetExists("DetallesControl") Then
        Set Asignaciones = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        Asignaciones.Name = "DetallesControl"
    Else
        Set Asignaciones = Worksheets("DetallesControl")
    End If
        
    Asignaciones.Activate
    Asignaciones.Range("A:Z").ColumnWidth = 30
    Asignaciones.Range("A:AZ").Clear
    UltimoRenglon = Worksheets("Ventas").Range("A" & Worksheets("Ventas").Rows.Count).End(xlUp).Row

    Worksheets("Ventas").Range("A1:Z" & UltimoRenglon).Copy
    Asignaciones.Range("AA1").PasteSpecial (xlPasteValues)

    Set RangeBase = Asignaciones.Range("A1:AZ" & UltimoRenglon)
    RangeBase.RemoveDuplicates Columns:=Array(27), Header:=xlNo
    
    UltimoRenglon = Asignaciones.Range("AA" & Asignaciones.Rows.Count).End(xlUp).Row
    If UltimoRenglon > 1 Then
        Asignaciones.Range("A2:A" & UltimoRenglon) = "Factura"
        
        Asignaciones.Range("AA2:AA" & UltimoRenglon).Copy
        Asignaciones.Range("B2:B" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Asignaciones.Range("AC2:AC" & UltimoRenglon).Copy
        Asignaciones.Range("C2:C" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Asignaciones.Range("AQ2:AQ" & UltimoRenglon).Copy
        Asignaciones.Range("D2:D" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Asignaciones.Range("E2:E" & UltimoRenglon) = "VENTA"
        
        Asignaciones.Range("F2:F" & UltimoRenglon).FormulaR1C1 = "=" & Chr(34) & "Venta            " & Chr(34) & "&TEXT(RC[31]," & Chr(34) & "DD-MMM-AAAA" & Chr(34) & ")"
        Asignaciones.Range("F2:F" & UltimoRenglon).Copy
        Asignaciones.Range("F2:F" & UltimoRenglon).PasteSpecial (xlPasteValues)
        Asignaciones.Range("F2:F" & UltimoRenglon).NumberFormat = "dd-mmm-yyyy"
        
        Asignaciones.Range("AK2:AK" & UltimoRenglon).Copy
        Asignaciones.Range("Z2:Z" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Asignaciones.Range("Y2:Y" & UltimoRenglon).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=SUMIF(Ventas!C[-24],DetallesControl!RC[2],Ventas!C[-8])"
        Asignaciones.Range("Y2:Y" & UltimoRenglon).Copy
        Asignaciones.Range("Y2:Y" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Asignaciones.Range("Y2:Y" & UltimoRenglon).Copy
        Asignaciones.Range("W2:W" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Asignaciones.Range("G2:G" & UltimoRenglon).FormulaR1C1 = "=formatomoneda(RC[18],Preferencias!R4C2,20)"
        
        
        Asignaciones.Range("G2:G" & UltimoRenglon).Copy
        Asignaciones.Range("G2:G" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Asignaciones.Range("H2:H" & UltimoRenglon).FormulaR1C1 = "=NOACENTOS(RC[-6]&" & Chr(34) & " " & Chr(34) & "&RC[-5]&" & Chr(34) & " " & Chr(34) & "&RC[-3]&" & Chr(34) & " " & Chr(34) & "&RC[-2])"
        Asignaciones.Range("H2:H" & UltimoRenglon).Copy
        Asignaciones.Range("H2:H" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
    End If
    If Not SheetExists("Temporal") Then
        Set Devoluciones = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        Devoluciones.Name = "Temporal"
    Else
        Set Devoluciones = Worksheets("Temporal")
    End If
    
    Devoluciones.Activate
    Devoluciones.Range("A:Z").ColumnWidth = 30
    Devoluciones.Range("A:AZ").Clear
    UltimoRenglon = Worksheets("Devoluciones").Range("A" & Worksheets("Devoluciones").Rows.Count).End(xlUp).Row
    Worksheets("Devoluciones").Range("A1:Z" & UltimoRenglon).Copy
    Devoluciones.Range("AA" & 1).PasteSpecial (xlPasteValues)
    
    UltimoRenglon = Devoluciones.Range("AA" & Devoluciones.Rows.Count).End(xlUp).Row
    Set RangeBase = Devoluciones.Range("A1:AZ" & UltimoRenglon)
    RangeBase.RemoveDuplicates Columns:=Array(27), Header:=xlNo   'Array(27, 47)
 
    UltimoRenglon = Devoluciones.Range("AA" & Devoluciones.Rows.Count).End(xlUp).Row
    If UltimoRenglon > 1 Then
    
        Devoluciones.Range("A2:A" & UltimoRenglon) = "xxxxx"
        
        Devoluciones.Range("AA2:AA" & UltimoRenglon).Copy    ' Devolucion
        Devoluciones.Range("B2:B" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Devoluciones.Range("AC2:AC" & UltimoRenglon).Copy
        Devoluciones.Range("C2:C" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Devoluciones.Range("AQ2:AQ" & UltimoRenglon).Copy
        Devoluciones.Range("D2:D" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Devoluciones.Range("E2:E" & UltimoRenglon) = "DEVOLUCION"
        
        Devoluciones.Range("F2:F" & UltimoRenglon).FormulaR1C1 = "=" & Chr(34) & "Devolución     " & Chr(34) & "&TEXT(RC[42]," & Chr(34) & "DD-MMM-AAAA" & Chr(34) & ")"
        Devoluciones.Range("F2:F" & UltimoRenglon).Copy
        Devoluciones.Range("F2:F" & UltimoRenglon).PasteSpecial (xlPasteValues)
        Devoluciones.Range("F2:F" & UltimoRenglon).NumberFormat = "dd-mmm-yyyy"
        
        Devoluciones.Range("AV2:AV" & UltimoRenglon).Copy
        Devoluciones.Range("Z2:Z" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Devoluciones.Range("Y2:Y" & UltimoRenglon).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=SUMIF(Devoluciones!C[-4],Temporal!RC[22],Devoluciones!C[-15])"
        Devoluciones.Range("Y2:Y" & UltimoRenglon).Copy
        Devoluciones.Range("Y2:Y" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Devoluciones.Range("Y2:Y" & UltimoRenglon).Copy
        Devoluciones.Range("X2:X" & UltimoRenglon).PasteSpecial (xlPasteValues)
        
        Devoluciones.Range("G2:G" & UltimoRenglon).FormulaR1C1 = "=formatomoneda(RC[18],Preferencias!R4C2,20)"
        Devoluciones.Range("G2:G" & UltimoRenglon).Copy
        Devoluciones.Range("G2:G" & UltimoRenglon).PasteSpecial (xlPasteValues)                                                                '=SUM(RC[36])    '18     '36
        
        Devoluciones.Range("H2:H" & UltimoRenglon).FormulaR1C1 = "=NOACENTOS(RC[-6]&" & Chr(34) & " " & Chr(34) & "&RC[-5]&" & Chr(34) & " " & Chr(34) & "&RC[-3]&" & Chr(34) & " " & Chr(34) & "&RC[-2])"
        Devoluciones.Range("H2:H" & UltimoRenglon).Copy
        Devoluciones.Range("H2:H" & UltimoRenglon).PasteSpecial (xlPasteValues)
    End If
    
    RenglonPegar = Asignaciones.Range("AA" & Asignaciones.Rows.Count).End(xlUp).Row + 1
    UltimoRenglon = Devoluciones.Range("AA" & Devoluciones.Rows.Count).End(xlUp).Row
    
    If UltimoRenglon > 1 Then
        Devoluciones.Range("A2:AZ" & UltimoRenglon).Copy
        Asignaciones.Range("A" & RenglonPegar).PasteSpecial (xlPasteValues)
    End If
    
    'Etiquetas
    Worksheets("Ventas").Range("B1:AZ1").Copy
    Asignaciones.Range("AB1").PasteSpecial
    
    'Asignaciones.Range("A1:AZ1").PasteSpecial (xlPasteValues)
    
    Asignaciones.Activate
    Columns("A:AZ").Sort key1:=Range("B1"), order1:=xlDescending, key2:=Range("E1"), order2:=xlDescending, Header:=xlYes
    
    UltimoRenglon = Asignaciones.Range("AA" & Asignaciones.Rows.Count).End(xlUp).Row
    If UltimoRenglon < 2 Then
        Exit Sub
    End If
    Asignaciones.Range("J2:J" & UltimoRenglon).FormulaR1C1 = "=ROW(RC[-1])"
    Asignaciones.Range("J2:J" & UltimoRenglon).Copy
    Asignaciones.Range("J2:J" & UltimoRenglon).PasteSpecial (xlPasteValues)
    
    UltimoRenglon = Asignaciones.Range("A" & Asignaciones.Rows.Count).End(xlUp).Row
    Set TheRange = Asignaciones.Range("A2:J" & UltimoRenglon)
    ArrayAsignaciones = TheRange
    
    ListBoxAsignaciones.List = ArrayAsignaciones
    
    If NombreInicio <> "" Then
        TextBoxTecnico = NombreInicio
        NombreInicio = ""
    End If
    
    Dim RangeSuma As Range
    Set RangeSuma = Worksheets("DetallesControl").Range("W:W")
    TextBoxSumaAsignaciones = Format(Application.WorksheetFunction.Sum(RangeSuma), "$ #,##0.00")
    TextBoxConteoAsignaciones = Application.WorksheetFunction.Count(RangeSuma)
    
    If ListBoxAsignaciones.ListCount > 0 Then
        ListBoxAsignaciones.Selected(0) = True
    End If
    
  


  
    
End Sub


Private Sub UserForm_Initialize()
    Set ArchivoPrincipal = ActiveWorkbook
    Inicio = True
      
    index = 1
    num = 0
End Sub

Private Sub UserForm_Terminate()
    ArchivoPrincipal.Activate
    'Application.Calculation = xlCalculationAutomatic
    'ThisWorkbook.Application.Visible = True
    'Application.ScreenUpdating = True
    'Application.DisplayAlerts = True


    'ThisWorkbook.Application.Visible = True
    NombreInicio = ""
End Sub



