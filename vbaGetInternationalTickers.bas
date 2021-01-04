Attribute VB_Name = "vbaGetInternationalTickers"
'              ````````````````````````````````````````````....................................`
'           ```````````````````````````````..........................................-----------.`
'          ````````````````````````````````..........................................-------------
'          ````````````````````````````````..........................................------------.
'         `````````````````````````````````..........................................------------`
'         ````````.----````....------..````.......-----..------........-------::::-..------------
'         ```````/+oo++.```sdddddddddddyo/.......oddddo..sddddd+.....-yddddy::+ooo+/------------.
'        ```````:ooo-..```.mmmmmmmmmmmmmmmh+.....dmmmm:...ymmmmmo...+dmmmmo-.---:ooo:-----------
'        ```````:ooo``````/mmmmh:::/+sdmmmmm+.../mmmmd.....ymmmmmo:ymmmmh:......-ooo:----------.
'        ```````:ooo``````ymmmm+``````-ymmmmd...smmmmo......ymmmmmdmmmd+-.......-ooo:----------`
'       ````````/oo+`````.mmmmm-```````/mmmmm...mmmmm:.......ymmmmmmms-.........-ooo:----------
'       ``````.:+oo:`````+mmmmh````````ommmmh../mmmmd.......-smmmmmmy............+oo+:--------.
'      ```````/+oo/.`````hmmmm+```````/mmmmm:..ymmmmo....../hmmmmmmmmo...........:+oo+/-------`
'      ````````.+oo/````-mmmmm.```..:smmmmd/...mmmmm-....-smmmmhymmmmmo.........-+oo/--------.
'      `````````:ooo````+mmmmmssyyhdmmmmmy:...+mmmmh....+dmmmmo--ymmmmmo........-ooo:--------`
'     ``````````:ooo````hmmmmmmmmmmmmmhs:.`...ymmmm+..:ymmmmh:....ymmmmmo.......-ooo:--------
'     ``````````:ooo```.yyyyyyyyyso+/-.````...yyyyy-.:syyyyo......-syyyyy:......-ooo:-------.
'    ```````````-ooo/-:.```````````````````..................................-::+ooo--------`
'    ````````````-/++++.```````````````````..................................:++++/-.------.
'    ```````````````..`````````````````````...................................----...------`
'   ```````````````````````````````````````..........................................------
'    ``````````````````````````````````````..........................................-----.
'      ``````````````````````````````````````............................................`
'                                                                 ````````````````````

Sub IE_1_BMV()
Dim IE As New InternetExplorer
Dim i, i2, i3, currPage, mxPags, newfila, x, y As Integer
Dim dataBMV(10, 13) As Variant
Dim h(13) As String

h(1) = "ISSUER"
h(2) = "SERIES"
h(3) = "TIME"
h(4) = "LAST"
h(5) = "VWAP"
h(6) = "PREVIOUS"
h(7) = "MAXIMUM"
h(8) = "MINIMUM"
h(9) = "VOLUME"
h(10) = "AMOUNT"
h(11) = "OPS."
h(12) = "Change Points"
h(13) = "Change %"

'Descarga de encabezados
For i = 1 To 13: ActiveSheet.Cells(1, i).Value = h(i): Next i

'inmovilizar panel
ActiveSheet.Cells(2, 2).Select
ActiveWindow.FreezePanes = True



'estilos encabezados
With ActiveSheet.Range("A1:M1")
  .Font.Bold = True
  .Font.Color = vbWhite
  .Interior.Color = RGB(37, 67, 103)
  
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlBottom
  .ReadingOrder = xlContext
End With

'navegar a la página de la bolsa mexicana de valores
With IE
  .navigate "https://www.bmv.com.mx/en/markets/global-market"
  .Visible = True
  .Top = 0
  .Left = -5
  .Height = 750
  .Width = 1950
End With
 
  'mientras la pagina de la bmv no este cargada se esperara
  Do While IE.readyState <> READYSTATE_COMPLETE Or IE.Busy = True: DoEvents: Loop


'seleccionar el mercado de capitales
  
  'seleccionar el combobox de capitales
  IE.document.getElementById("mglobalCB1").getElementsByClassName("value")(0).Children(1).Click
  
  'bucle encontrar el valor de mercado de capitales y seleccionarlo
  i = 0
  Do While i <> IE.document.getElementById("mglobalCB1").getElementsByTagName("ul")(0).getElementsByTagName("li").Length
    If Trim(IE.document.getElementById("mglobalCB1").getElementsByTagName("ul")(0).getElementsByTagName("li")(i).Children(0).innerText) = "SIC Capitales" Then
      IE.document.getElementById("mglobalCB1").getElementsByTagName("ul")(0).getElementsByTagName("li")(i).Children(0).Click
      Exit Do
    End If
    i = i + 1
  Loop

  'Seleccionar el combobox de categoria
  IE.document.getElementById("mglobalCB2").getElementsByClassName("value")(0).Children(1).Click
  
  'bucle para seleccionar el concepto de "series  operadaas"
  i = 0
  Do While i2 <> IE.document.getElementById("mglobalCB2").getElementsByTagName("ul")(0).getElementsByTagName("li").Length
    If Trim(IE.document.getElementById("mglobalCB2").getElementsByTagName("ul")(0).getElementsByTagName("li")(i).Children(0).innerText) = "Series Operadas" Then
      IE.document.getElementById("mglobalCB2").getElementsByTagName("ul")(0).getElementsByTagName("li")(i).Children(0).Click
      Exit Do
    End If
    i = i + 1
  Loop
  
  'dar click en el botón de buscar
  IE.document.getElementsByClassName("btn")(0).Click
  
'Numero de repeticiones que se debe de hacer la tarea de descarga
mxPags = Val(Trim(IE.document.getElementById("tableCaoOp_paginate").getElementsByTagName("span")(0).getElementsByTagName("a")(5).innerText)) + 1
'Inicio del bucle
i2 = 1
Do While i2 < mxPags
  On Error Resume Next
  Debug.Print i2 & " / " & mxPags & " -> " & WorksheetFunction.Round((i2 / mxPags) * 100, 2) & " % | " & TimeValue(Now)
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'bucle obteneción de los datos y guardados en el arreglo
    i = 0
    Do While i <> IE.document.getElementById("tableCaoOp").getElementsByTagName("tbody")(0).getElementsByTagName("tr").Length
      
      With IE.document.getElementById("tableCaoOp").getElementsByTagName("tbody")(0).getElementsByTagName("tr")(i) 'cambiar el ultimo cero por el numero de fila
        For i3 = 0 To 12
        'ISSUER (0) | SERIES (1) | TIME (2) | LAST (3) | VWAP (4) | PREVIOUS (5) | MAXIMUM (6) | MINIMUM (7) | VOLUME (8) | AMOUNT (9) | OPS. (10) | Change Points (11) | Change % (12)
          dataBMV(i + 1, i3 + 1) = .getElementsByTagName("td")(i3).innerText
        Next i3
      End With
      
      i = i + 1
    Loop
    
    'numero de fila a partir de la cual debe empezar a ejecutarse
    newfila = WorksheetFunction.CountA(Columns("A")) + 1
    
    'vaciado de la información
    On Error Resume Next
    For x = 0 To 9
      For y = 1 To 13
        ActiveSheet.Cells(x + newfila, y).Value = dataBMV(1 + x, y)
      Next y
    Next x
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  IE.document.getElementById("tableCaoOp_next").Click
  i2 = i2 + 1
Loop

IE.Quit

'Eliminar los duplicados
ActiveSheet.Cells(1, 1).CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), Header:=xlYes

Debug.Print "Descarga finalizada"

End Sub
