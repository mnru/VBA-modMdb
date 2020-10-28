VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMdbInfo 
   Caption         =   "mdbInfo"
   ClientHeight    =   2805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
   OleObjectBlob   =   "frmMdbInfo.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMdbInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Dim cat 'ADOX.Catalog
Dim curTblName
Dim curTblType

Private Sub cmdEdit_Click()
    frmSQLEdit.Show
End Sub

Private Sub cmdCompact_Click()
    Call compactDB(Me.tbxPath.Value)
End Sub

Private Sub cmdList_Click()
    Call setCurrentTbl
    If curTblName = "" Then Exit Sub
    curR = ActiveCell.Row
    curC = ActiveCell.Column
    Select Case curTblType
        Case "Table"
            ary = mkTblColAry
            Call layAryAt(ary, curR, curC, "c")
        Case "View"
            ary = mkViewColAry
            Call layAryAt(ary, curR, curC, "c")
        Case Else
    End Select
End Sub

Private Sub cmdRefresh_Click()
    Call refreshCatData
End Sub

Private Sub cmdSQL_Click()
    formDBPath = Me.tbxPath
    frmSQLEdit.Show
End Sub

Private Sub UserForm_Initialize()
    Set cat = CreateObject("ADOX.Catalog")
    Set fso = CreateObject("Scripting.FileSystemObject")
    spath = Application.GetOpenFilename("database file,*.mdb;*.accdb", , "Select mdb file.")
    If TypeName(spath) = "Boolean" Then Exit Sub
    Me.tbxPath = spath
    Me.tbxPath.Enabled = False
    ext = fso.GetExtensionName(spath)
    bmdb = (ext = "mdb")
    Call refreshCatData
End Sub

Sub refreshCatData()
    cat.activeConnection = mdbConInfo(Me.tbxPath.Value)
    cboTable.Clear
    cboView.Clear
    cboLink.Clear
    cboProcedure.Clear
    For Each tbl In cat.tables
        If Left(tbl.name, 4) <> "MSys" Then
            Select Case tbl.Type
                Case "TABLE": cboTable.AddItem tbl.name
                Case "VIEW": cboView.AddItem tbl.name
                Case "LINK": cboLink.AddItem tbl.name
                Case Else
            End Select
        End If
    Next tbl
    For Each prc In cat.Procedures
        cboProcedure.AddItem prc.name
    Next prc
    OptTable.Value = True
    Call OptTable_Click
    curTblName = ""
    curTblType = "Table"
End Sub

Private Sub OptTable_Click()
    cboTable.Visible = True
    cboView.Visible = False
    cboLink.Visible = False
    cboProcedure.Visible = False
End Sub

Private Sub OptView_Click()
    cboTable.Visible = False
    cboView.Visible = True
    cboLink.Visible = False
    cboProcedure.Visible = False
End Sub

Private Sub OptLink_Click()
    cboTable.Visible = False
    cboView.Visible = False
    cboLink.Visible = True
    cboProcedure.Visible = False
End Sub

Private Sub OptProcedure_Click()
    cboTable.Visible = False
    cboView.Visible = False
    cboLink.Visible = False
    cboProcedure.Visible = True
End Sub

Sub setCurrentTbl()
    curTblName = ""
    If OptTable Then
        curTblName = cboTable.text
        curTblType = "Table"
    ElseIf OptView Then
        curTblName = cboView.text
        curTblType = "View"
    ElseIf OptLink Then
        curTblName = cboLink.text
        curTblType = "Link"
    ElseIf OptProcedure Then
        curTblName = cboProcedure.text
        curTblType = "Procedure"
    End If
End Sub

Private Sub cmdData_Click()
    Call setCurrentTbl
    If curTblName <> "" Then
        Call displayQueryTable(curTblName, , tbxPath.Value, , curTblType)
    End If
End Sub

Private Sub cmdDef_Click()
    Call setCurrentTbl
    Select Case curTblType
        Case "Table"
            Call mkTblDefSheet("A4")
        Case "View"
            Call mkViewDefSheet("A4")
        Case "Link"
            Call mkLinkDefSheet("A4")
        Case "Procedure"
            Call mkProcDefSheet("A4")
        Case Else
    End Select
End Sub

Sub mkViewDefSheet(sTbl, Optional top = "A4")
    Dim curR, curC
    Dim sn, sSQL
    If curTblName = "" Then Exit Sub
    '
    Sheets.Add After:=Sheets(Sheets.Count)
    sn = ActiveSheet.name
    curR = Range(top).Row
    curC = Range(top).Column
    '
    sSQL = cat.Views(curTblName).Command.CommandText
    Call layAryAt(Array("Path", tbxPath.Value), curR - 3, curC)
    Call layAryAt(Array(curTblType, curTblName), curR - 2, curC)
    Call layAryAt(Array("SQL", sSQL), curR, curC)
End Sub

Sub mkProcDefSheet(Optional top = "A4")
    Dim curR, curC
    Dim sn, sSQL
    If curTblName = "" Then Exit Sub
    '
    Sheets.Add After:=Sheets(Sheets.Count)
    sn = ActiveSheet.name
    curR = Range(top).Row
    curC = Range(top).Column
    '
    sSQL = cat.Procedures(curTblName).Command.CommandText
    Call layAryAt(Array("Path", tbxPath.Value), curR - 3, curC)
    Call layAryAt(Array(curTblType, curTblName), curR - 2, curC)
    Call layAryAt(Array("SQL", sSQL), curR, curC)
End Sub

Sub mkLinkDefSheet(Optional top = "A4")
    Dim curR, curC
    Dim sn, sSQL, toPath, toTable
    Dim title, data
    If curTblName = "" Then Exit Sub
    '
    Sheets.Add After:=Sheets(Sheets.Count)
    sn = ActiveSheet.name
    curR = Range(top).Row
    curC = Range(top).Column
    '
    toPath = cat.tables(curTblName).Properties("Jet OLEDB:Link Datasource")
    toTable = cat.tables(curTblName).Properties("Jet OLEDB:Remote Table Name")
    '
    title = Array("link from(virtual)", "path", "table", "link to(real)", "path", "table")
    data = Array("", tbxPath.Value, curTblName, "", toPath, toTable)
    Call layAryAt(title, curR, curC, "c")
    Call layAryAt(data, curR, curC + 1, "c")
End Sub

Function mkTblColAry()
    Dim num, ret, obj
    Set obj = cat.tables(curTblName)
    num = obj.Columns.Count
    ReDim ret(0 To num)
    ret(0) = curTblName
    For i = 1 To num
        ret(i) = obj.Columns(i - 1).name
    Next i
    mkTblColAry = ret
    Set obj = Nothing
End Function

Function mkViewColAry()
    Dim num, ret, obj
    Set obj = cat.Views(curTblName)
    num = obj.Columns.Count
    ReDim ret(0 To num)
    ret(0) = curTblName
    For i = 1 To num
        ret(i) = obj.Columns(i - 1).name
    Next i
    mkTblColAry = ret
    Set obj = Nothing
End Function

Sub mkTblDefSheet(Optional top = "A4")
    Dim colNum, keyNum, i, eType, eAtr, curR, curC, sz
    Dim sn, sname
    Dim tblObj
    If curTblName = "" Then Exit Sub
    '
    Sheets.Add After:=Sheets(Sheets.Count)
    sn = ActiveSheet.name
    curR = Range(top).Row
    curC = Range(top).Column
    Call layAryAt(Array("Path", tbxPath.Value), curR - 3, curC)
    Call layAryAt(Array(curTblType, curTblName), curR - 2, curC)
    '
    title = Array("name", "type", "fixed Length", "not null", "max size")
    Call layAryAt(title, curR, curC)
    curR = curR + 1
    '
    Set tblObj = cat.tables(curTblName)
    colNum = tblObj.Columns.Count
    ReDim clmAry(1 To colNum)
    ReDim enumAry(1 To colNum)
    i = 1
    For Each clm In tblObj.Columns
        sname = clm.name
        eType = clm.Type
        eAtr = clm.Attributes '1:adColFixed,2:adColNullable
        sz = clm.definedsize
        clmAry(i) = sname
        dataAry = Array(sname, getEnumColumnTypeName(eType), IIf((eAtr And 1) = 1, "〇", ""), IIf((eAtr And 2) = 2, "", "〇"), IIf(sz = 0, "", sz))
        Call layAryAt(dataAry, curR, curC)
        curR = curR + 1
        i = i + 1
    Next clm
    'properties  default,autoincrement
    curR = Range(top).Row + 1
    curC = Range(top).Column + 5
    idxn = 1
    bolAry = Array(True, False)
    For Each bolval In bolAry
        For Each idx In tblObj.indexes
            If idx.primarykey = bolval Then
                Cells(curR - 1, curC) = IIf(idx.primarykey, "primary key", "index" & idxn)
                i = 1
                ReDim idxAry(1 To colNum)
                For Each elm In idx.Columns
                    n = Application.WorksheetFunction.Match(elm.name, clmAry)
                    idxAry(n) = i
                    i = i + 1
                Next
                Call layAryAt(idxAry, curR, curC, "c")
                curC = curC + 1
                If Not idx.primarykey Then idxn = idxn + 1
            End If
        Next
    Next
End Sub

Private Sub layAryAt(ary, r, c, Optional rc = "r", Optional sn = "", Optional bn = "")
    If sn = "" Then sn = ActiveSheet.name
    If bn = "" Then bn = ActiveWorkbook.name
    n = lenAry(ary)
    Select Case rc
        Case "r"
            Workbooks(bn).Worksheets(sn).Cells(r, c).Resize(1, n) = ary
        Case "c"
            Workbooks(bn).Worksheets(sn).Cells(r, c).Resize(n, 1) = Application.WorksheetFunction.Transpose(ary)
        Case Else
    End Select
End Sub

Private Function lenAry(ary As Variant, Optional dm = 1) As Long
    lenAry = UBound(ary, dm) - LBound(ary, dm) + 1
End Function

Function getEnumColumnTypeName(num)
    Dim ret
    Select Case num
        Case 0: ret = "adEmpty"
        Case 2: ret = "adSmallInt"
        Case 3: ret = "adInteger"
        Case 4: ret = "adSingle"
        Case 5: ret = "adDouble"
        Case 6: ret = "adCurrency"
        Case 7: ret = "adDate"
        Case 8: ret = "adBSTR"
        Case 9: ret = "adIDispatch"
        Case 10: ret = "adError"
        Case 11: ret = "adBoolean"
        Case 12: ret = "adVariant"
        Case 13: ret = "adIUnknown"
        Case 14: ret = "adDecimal"
        Case 16: ret = "adTinyInt"
        Case 17: ret = "adUnsignedTinyInt"
        Case 18: ret = "adUnsignedSmallInt"
        Case 19: ret = "adUnsignedInt"
        Case 20: ret = "adBigInt"
        Case 21: ret = "adUnsignedBigInt"
        Case 64: ret = "adFileTime"
        Case 72: ret = "adGUID"
        Case 128: ret = "adBinary"
        Case 129: ret = "adChar"
        Case 130: ret = "adWChar"
        Case 131: ret = "adNumeric"
        Case 132: ret = "adUserDefined"
        Case 133: ret = "adDBDate"
        Case 134: ret = "adDBTime"
        Case 135: ret = "adDBTimeStamp"
        Case 136: ret = "adChapter"
        Case 138: ret = "adPropVariant"
        Case 139: ret = "adVarNumeric"
        Case 200: ret = "adVarChar"
        Case 201: ret = "adLongVarChar"
        Case 202: ret = "adVarWChar"
        Case 203: ret = "adLongVarWChar"
        Case 204: ret = "adVarBinary"
        Case 205: ret = "adLongVarBinary"
        Case Else: ret = ""
    End Select
    getEnumColumnTypeName = ret
End Function

Function getColumnTypeSize(name)
    Dim ret
    Select Case name
        Case "adBoolean"
            ret = "1bit"
        Case "adTinyInt", "adUnsignedTinyInt"
            ret = "1byte"
        Case "adSmallInt", "adUnsignedSmallInt"
            ret = "2byte"
        Case "adInteger", "adSingle", "adUnsignedInt"
            ret = "4byte"
        Case "adDouble", "adCurrency", "adDate", "adBigInt", "adUnsignedInt"
            ret = "8byte"
        Case "adDecimal"
            ret = "12byte"
        Case Else
            ret = ""
    End Select
    getColumnTypeSize = ret
End Function
