
Attribute VB_Name = "VITAC_JSON"
Option Explicit

' Requires reference: Microsoft Scripting Runtime (for Dictionary)
' Tools > References... > check "Microsoft Scripting Runtime"
' Or use late binding on Dictionary (as written below).

Public Sub ExportVITACJSON()
    Dim wb As Workbook
    Dim shBase As Worksheet, shMeta As Worksheet
    On Error Resume Next
    Set wb = ThisWorkbook
    Set shBase = wb.Worksheets("BASE_DATOS")
    Set shMeta = wb.Worksheets("LANCES_CAPTURAS")
    On Error GoTo 0
    If shBase Is Nothing Or shMeta Is Nothing Then
        MsgBox "Faltan hojas BASE_DATOS o LANCES_CAPTURAS.", vbExclamation
        Exit Sub
    End If

    Dim lastRowB As Long, lastRowM As Long
    lastRowB = shBase.Cells(shBase.Rows.Count, 1).End(xlUp).Row
    lastRowM = shMeta.Cells(shMeta.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    Dim especieCol As Long, lanceCol As Long, tallaCol As Long
    especieCol = GetCol(shBase, "Especie")
    lanceCol = GetCol(shBase, "Lance")
    tallaCol = GetCol(shBase, "Talla")
    If especieCol = 0 Or lanceCol = 0 Or tallaCol = 0 Then
        MsgBox "No se encontraron columnas requeridas en BASE_DATOS.", vbExclamation
        Exit Sub
    End If

    Dim metaCols As Object: Set metaCols = CreateObject("Scripting.Dictionary")
    Dim headersM As Variant: headersM = GetHeaders(shMeta)
    For i = LBound(headersM) To UBound(headersM)
        metaCols(headersM(i)) = i + 1
    Next i

    If Not metaCols.Exists("Lance") Then
        MsgBox "No se encontró columna Lance en LANCES_CAPTURAS.", vbExclamation
        Exit Sub
    End If

    ' Collect lances
    Dim lancesDict As Object: Set lancesDict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRowM
        If shMeta.Cells(i, metaCols("Lance")).Value <> "" Then
            lancesDict(CLng(shMeta.Cells(i, metaCols("Lance")).Value)) = True
        End If
    Next i

    Dim lances() As Long, k As Long
    ReDim lances(0 To lancesDict.Count - 1)
    k = 0
    Dim key As Variant
    For Each key In lancesDict.Keys
        lances(k) = CLng(key)
        k = k + 1
    Next key
    QuickSortLong lances, LBound(lances), UBound(lances)

    ' Determine bins from data
    Dim tmin As Long, tmax As Long
    tmin = 999999: tmax = -999999
    For i = 2 To lastRowB
        If shBase.Cells(i, tallaCol).Value <> "" Then
            Dim t As Double: t = CDbl(shBase.Cells(i, tallaCol).Value)
            If t < tmin Then tmin = Int(t / 5) * 5
            If t > tmax Then tmax = -Int(-t / 5) * 5 ' ceil
        End If
    Next i
    If tmin > tmax Then
        tmin = 15: tmax = 125
    End If
    If tmax <= tmin Then tmax = tmin + 5

    Dim labels As Object: Set labels = CreateObject("Scripting.Dictionary")
    Dim a As Long
    For a = tmin To tmax - 1 Step 5
        labels.Add CStr(a) & "-" & CStr(a + 4), True
    Next a

    ' Frequency by lance/species
    Dim dataByLance As Object: Set dataByLance = CreateObject("Scripting.Dictionary")
    Dim dataMsur As Object: Set dataMsur = CreateObject("Scripting.Dictionary")

    Dim spec As String, l As Long, labelKey As String, idx As Long
    Dim arrCola() As Long, arrSur() As Long
    ReDim arrCola(0 To labels.Count - 1)
    ReDim arrSur(0 To labels.Count - 1)

    ' Initialize per-lance arrays
    For Each key In lances
        dataByLance(CStr(key)) = ZerosArray(labels.Count)
        dataMsur(CStr(key)) = ZerosArray(labels.Count)
    Next key

    ' Build an index of label lower bounds to position
    Dim labelIdx As Object: Set labelIdx = CreateObject("Scripting.Dictionary")
    idx = 0
    For Each key In labels.Keys
        Dim low As Long: low = CLng(Split(key, "-")(0))
        labelIdx(CStr(low)) = idx
        idx = idx + 1
    Next key

    ' Fill frequencies
    For i = 2 To lastRowB
        spec = LCase$(Trim$(CStr(shBase.Cells(i, especieCol).Value)))
        If spec <> "" Then
            l = CLng(shBase.Cells(i, lanceCol).Value)
            Dim talla As Double: talla = CDbl(shBase.Cells(i, tallaCol).Value)
            Dim lowBound As Long: lowBound = Int(talla / 5) * 5
            If lowBound < tmin Then lowBound = tmin
            If lowBound > tmax - 5 Then lowBound = tmax - 5
            If labelIdx.Exists(CStr(lowBound)) Then
                idx = CLng(labelIdx(CStr(lowBound)))
                If InStr(spec, "mcola") > 0 Then
                    Dim arr1 As Variant: arr1 = dataByLance(CStr(l))
                    arr1(idx) = CLng(arr1(idx)) + 1
                    dataByLance(CStr(l)) = arr1
                ElseIf InStr(spec, "msur") > 0 Or InStr(spec, "sur") > 0 Then
                    Dim arr2 As Variant: arr2 = dataMsur(CStr(l))
                    arr2(idx) = CLng(arr2(idx)) + 1
                    dataMsur(CStr(l)) = arr2
                End If
            End If
        End If
    Next i

    ' dataCapt, lanceInfo, coordsLance
    Dim dataCapt As Object: Set dataCapt = CreateObject("Scripting.Dictionary")
    Dim lanceInfo As Object: Set lanceInfo = CreateObject("Scripting.Dictionary")
    Dim coordsLance As Object: Set coordsLance = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRowM
        If shMeta.Cells(i, metaCols("Lance")).Value = "" Then GoTo NextRow
        l = CLng(shMeta.Cells(i, metaCols("Lance")).Value)

        Dim msw As Double: msw = Nz(shMeta.Cells(i, GetColBy(metaCols, "MsurW")).Value)
        Dim mcw As Double: mcw = Nz(shMeta.Cells(i, GetColBy(metaCols, "McolaW")).Value)
        Dim ow As Double:  ow  = Nz(shMeta.Cells(i, GetColBy(metaCols, "OtrosW")).Value)

        Dim mp As Variant: mp = shMeta.Cells(i, GetColBy(metaCols, "Msur%")).Value
        Dim mcp As Variant: mcp = shMeta.Cells(i, GetColBy(metaCols, "Mcola%")).Value
        Dim op As Variant:  op  = shMeta.Cells(i, GetColBy(metaCols, "Otros%")).Value

        If IsEmpty(mp) Or mp = "" Or IsEmpty(mcp) Or mcp = "" Or IsEmpty(op) Or op = "" Then
            Dim tot As Double: tot = msw + mcw + ow
            If tot > 0 Then
                mp = 100# * msw / tot
                mcp = 100# * mcw / tot
                op = 100# * ow / tot
            Else
                mp = 0: mcp = 0: op = 0
            End If
        End If

        Dim c As Object: Set c = CreateObject("Scripting.Dictionary")
        c("Msur") = Round(CDbl(mp), 2)
        c("Mcola") = Round(CDbl(mcp), 2)
        c("Otros") = Round(CDbl(op), 2)
        dataCapt(CStr(l)) = c

        ' Coords
        Dim lat As Variant: lat = shMeta.Cells(i, GetColBy(metaCols, "Latitud1")).Value
        Dim lon As Variant: lon = shMeta.Cells(i, GetColBy(metaCols, "Longitud1")).Value
        If Not IsEmpty(lat) And Not IsEmpty(lon) And lat <> "" And lon <> "" Then
            Dim arr As Variant: arr = Array(CDbl(lat), CDbl(lon))
            coordsLance(CStr(l)) = arr
        End If

        Dim fecha As String, hora As String
        fecha = CStr(shMeta.Cells(i, GetColBy(metaCols, "Fecha")).Text)
        hora = CStr(shMeta.Cells(i, GetColBy(metaCols, "Hora")).Text)

        Dim info As Object: Set info = CreateObject("Scripting.Dictionary")
        info("fecha") = Trim$(fecha & " " & hora)
        info("latTxt") = FmtDeg(lat, True)
        info("lonTxt") = FmtDeg(lon, False)

        Dim kg As Object: Set kg = CreateObject("Scripting.Dictionary")
        kg("Msur") = msw: kg("Mcola") = mcw: kg("Otros") = ow
        info("kg") = kg
        lanceInfo(CStr(l)) = info

NextRow:
    Next i

    ' Build JSON string
    Dim json As String
    json = "{""classes"":" & JsonArrayLabels(labels) & "," & _
           """lances"":" & JsonArrayLong(lances) & "," & _
           """dataByLance"":" & JsonDictArrays(dataByLance) & "," & _
           """dataMsur"":" & JsonDictArrays(dataMsur) & "," & _
           """dataCapt"":" & JsonDictDict(dataCapt) & "," & _
           """lanceInfo"":" & JsonDictDict(lanceInfo) & "," & _
           """coordsLance"":" & JsonDictArrays(coordsLance) & "}"

    Dim outPath As String
    outPath = wb.Path & Application.PathSeparator & "data.json"
    Dim f As Integer: f = FreeFile
    Open outPath For Output As #f
    Print #f, json
    Close #f

    MsgBox "data.json exportado en:" & vbCrLf & outPath, vbInformation
End Sub

Private Function Nz(v As Variant, Optional d As Double = 0#) As Double
    If IsEmpty(v) Or v = "" Then Nz = d Else Nz = CDbl(v)
End Function

Private Function GetCol(ws As Worksheet, title As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim j As Long
    For j = 1 To lastCol
        If Trim$(LCase$(CStr(ws.Cells(1, j).Value))) = Trim$(LCase$(title)) Then
            GetCol = j: Exit Function
        End If
    Next j
    GetCol = 0
End Function

Private Function GetColBy(dict As Object, key As String) As Long
    If dict.Exists(key) Then GetColBy = dict(key) Else GetColBy = 0
End Function

Private Function GetHeaders(ws As Worksheet) As Variant
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim arr() As String: ReDim arr(1 To lastCol)
    Dim j As Long
    For j = 1 To lastCol
        arr(j) = CStr(ws.Cells(1, j).Value)
    Next j
    GetHeaders = arr
End Function

Private Function ZerosArray(n As Long) As Variant
    Dim arr() As Long: ReDim arr(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1: arr(i) = 0: Next i
    ZerosArray = arr
End Function

Private Function JsonArrayLong(arr() As Long) As String
    Dim s As String, i As Long
    s = "["
    For i = LBound(arr) To UBound(arr)
        s = s & CStr(arr(i))
        If i < UBound(arr) Then s = s & ","
    Next i
    s = s & "]"
    JsonArrayLong = s
End Function

Private Function JsonArrayLabels(labels As Object) As String
    Dim s As String, k As Variant, i As Long
    s = "["
    i = 0
    For Each k In labels.Keys
        s = s & """" & CStr(k) & """"
        i = i + 1
        If i < labels.Count Then s = s & ","
    Next k
    s = s & "]"
    JsonArrayLabels = s
End Function

Private Function JsonDictArrays(d As Object) As String
    Dim s As String, k As Variant
    s = "{"
    Dim first As Boolean: first = True
    For Each k In d.Keys
        If Not first Then s = s & "," Else first = False
        s = s & """" & CStr(k) & """:" & JsonArrayVariant(d(k))
    Next k
    s = s & "}"
    JsonDictArrays = s
End Function

Private Function JsonArrayVariant(v As Variant) As String
    Dim s As String: s = "["
    Dim i As Long
    For i = LBound(v) To UBound(v)
        If IsArray(v) Then
            s = s & CStr(v(i))
        Else
            s = s & CStr(v(i))
        End If
        If i < UBound(v) Then s = s & ","
    Next i
    s = s & "]"
    JsonArrayVariant = s
End Function

Private Function JsonDictDict(d As Object) As String
    Dim s As String, k As Variant
    s = "{"
    Dim first As Boolean: first = True
    For Each k In d.Keys
        If Not first Then s = s & "," Else first = False
        s = s & """" & CStr(k) & """:" & JsonObject(d(k))
    Next k
    s = s & "}"
    JsonDictDict = s
End Function

Private Function JsonObject(d As Object) As String
    Dim s As String, k As Variant
    s = "{"
    Dim first As Boolean: first = True
    For Each k In d.Keys
        If Not first Then s = s & "," Else first = False
        s = s & """" & CStr(k) & """:" & JsonValue(d(k))
    Next k
    s = s & "}"
    JsonObject = s
End Function

Private Function JsonValue(v As Variant) As String
    If IsObject(v) Then
        JsonValue = JsonObject(v)
    ElseIf IsArray(v) Then
        JsonValue = JsonArrayVariant(v)
    ElseIf IsNumeric(v) Then
        JsonValue = CStr(v)
    Else
        JsonValue = """" & Replace(CStr(v), """", "\""") & """"
    End If
End Function

Private Function FmtDeg(v As Variant, ByVal isLat As Boolean) As String
    On Error GoTo fail
    Dim val As Double: val = CDbl(v)
    Dim hemi As String
    If isLat Then
        hemi = IIf(val >= 0, "N", "S")
    Else
        hemi = IIf(val >= 0, "E", "W")
    End If
    FmtDeg = Format(Abs(val), "0.0000") & "° " & hemi
    Exit Function
fail:
    FmtDeg = ""
End Function

Private Sub QuickSortLong(a() As Long, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, x As Long, y As Long
    i = first: j = last
    x = a((first + last) \ 2)
    Do While i <= j
        Do While a(i) < x: i = i + 1: Loop
        Do While a(j) > x: j = j - 1: Loop
        If i <= j Then
            y = a(i): a(i) = a(j): a(j) = y
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortLong a, first, j
    If i < last Then QuickSortLong a, i, last
End Sub
