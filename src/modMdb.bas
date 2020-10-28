Attribute VB_Name = "modMdb"
Option Base 0
Public adoCon
Public formDBPath
Public curMDB
Const bACE = True

Function getCurMdb(Optional mdbPath = "", Optional bOverWrite = False, Optional bGlobal = True)
    If mdbPath = "" Then mdbPath = ThisWorkbook.path & "\data.mdb"
    If bGlobal And (curMDB = "" Or bOverWrite) Then curMDB = mdbPath
    ret = IIf(bGlobal, curMDB, mdbPath)
    getCurMdb = ret
End Function

Function setCurMdb(Optional mdbPath = "", Optional bOverWrite = False, Optional bGlobal = True)
    setCurMdb = getCurMdb(mdbPath, bOverWrite, bGlobal)
End Function

Function mdbConInfo(Optional spath = "")
    Dim cat, dic, fso
    Dim info As String, ext
    Set dic = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = fso.GetExtensionName(spath)
    '
    dic("Provider") = IIf(bACE, "Microsoft.ACE.OLEDB.12.0", "Microsoft.JET.OLEDB.4.0")
    dic("Jet OLEDB:Engine Type") = IIf(LCase(ext) = "mdb", 5, 6)
    dic("Data Source") = getCurMdb(spath)
    '
    mdbConInfo = dicToStr(dic, "=", ";")
End Function

Function mdbODBCInfo(Optional spath = "")
    Dim cat, dic, fso
    Dim info As String, ext
    Set dic = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = fso.GetExtensionName(spath)
    '
    dic("DSN") = "MS Access Database"
    dic("DBQ") = getCurMdb(spath)
    '
    mdbODBCInfo = dicToStr(dic, "=", ";")
End Function

Sub mkMdb(Optional spath = "", Optional bOverWrite = True)
    Dim dic, cat
    Dim info
    Set fso = CreateObject("Scripting.FileSystemObject")
    spath = getCurMdb(spath, False, False)
    ext = fso.GetExtensionName(spath)
    Set cat = CreateObject("ADOX.Catalog")
    info = mdbConInfo(spath)
    If bOverWrite Then
        If fso.fileexists(spath) Then fso.deletefile (spath)
    End If
    Call cat.Create(info)
End Sub

Sub mkMdbDialog()
    Dim path
    path = Application.GetSaveAsFilename("data", "mdb file,*.mdb,accdb file,*.accdb", 1, "select access file")
    If path = False Then Exit Sub
    Call mkMdb(path)
End Sub

Sub openMdbCon(Optional spath = "")
    info = mdbConInfo(getCurMdb(spath))
    ' Debug.Print info
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.Open (info)
    Set dic = Nothing
End Sub

Sub closeCon()
    adoCon.Close
    Set adoCon = Nothing
End Sub

Sub execSQL(sql, Optional spath = "")
    Call openMdbCon(getCurMdb(spath))
    On Error GoTo errorDispose
    adoCon.BeginTrans
    adoCon.Execute (sql)
    adoCon.CommitTrans
    adoCon.Close
    Exit Sub
errorDispose:
    MsgBox "rollback"
    adoCon.rollbacktrans
    adoCon.Close
    On Error GoTo 0
End Sub

Sub execSQLs(sqls, Optional spath = "", Optional bTrans = False)
    If bTrans Then
        Call execSQLsWithTransaction(sqls, spath)
        Exit Sub
    End If
    On Error GoTo errorDispose
    Call openMdbCon(getCurMdb(spath))
    For Each sql In sqls
        adoCon.BeginTrans
        Call adoCon.Execute(sql)
        adoCon.CommitTrans
    Next sql
    adoCon.Close
    Exit Sub
errorDispose:
    MsgBox "rollback"
    adoCon.rollbacktrans
    adoCon.Close
    On Error GoTo 0
End Sub

Sub execSQLsWithTransaction(sqls, Optional spath = "")
    Call openMdbCon(getCurMdb(spath))
    On Error GoTo errorDispose
    adoCon.BeginTrans
    For Each sql In sqls
        Call adoCon.Execute(sql)
    Next sql
    adoCon.CommitTrans
    adoCon.Close
    Exit Sub
errorDispose:
    MsgBox "rollback"
    adoCon.rollbacktrans
    adoCon.Close
    On Error GoTo 0
End Sub

Function mdbAsTable(spath, Optional tbl)
    ret = "[OLEDB;Database=" & spath & "].[" & tbl & "]"
    mdbAsTable = ret
End Function

Function txtAsTable(spath, Optional hdr = "")
    Set fso = CreateObject("Scripting.FileSystemObject")
    fn = fso.getfilename(spath)
    fdn = fso.GetParentFolderName(spath)
    If hdr <> "" Then hdr = ";HDR=" & hdr
    ret = "[TEXT;Database=" & fdn & hdr & "].[" & fn & "]"
    txtAsTable = ret
End Function

Function xlAsTable(spath, Optional sht, Optional rg = "", Optional bHdr = True)
    Dim hdr, ext
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = fso.GetExtensionName(spath)
    hdr = IIf(bHdr, "YES", "No")
    ret = "[" & excelDBType(ext) & ";Database=" & spath & ";HDR=" & hdr & "].[" & sht & "$" & rg & "]"
    xlAsTable = ret
End Function

Function excelDBType(ext)
    Dim ret
    Select Case LCase(ext)
        Case "xls": ret = "Excel 8.0"
        Case "xlsb": ret = "Excel 12.0"
        Case "xlsx": ret = "Excel 12.0 Xml"
        Case "xlsm": ret = "Excel 12.0 Macro"
        Case Else: ret = ""
    End Select
    excelDBType = ret
End Function

Function xlTblAsTable(xlTbl, Optional bkn = "", Optional bPlusAbove = True, Optional bHdr = True)
    Dim xlPath, adn, shn
    Dim rs, cs
    If bkn = "" Then bkn = ThisWorkbook.name
    Set fso = CreateObject("Scripting.FileSystemObject")
    Workbooks(bkn).Activate
    xlPath = Workbooks(bkn).FullName
    shn = Range(xlTbl).Parent.name
    Sheets(shn).Activate
    xlTbl = Trim(xlTbl)
    Set rg = Range(xlTbl)
    rs = rg.Rows.Count
    cs = rg.Columns.Count
    If bPlusAbove Then
        adn = rg.offset(-1, 0).Resize(rs + 1, cs).Address(False, False)
    Else
        adn = rg.Address(False, False)
    End If
    Set rg = Nothing
    xlTblAsTable = xlAsTable(xlPath, shn, adn, bHdr)
End Function

Function mkIniHead(Optional bHdr = True, Optional charset = "sjis", Optional dlm = "csv")
    Set dic = CreateObject("Scripting.Dictionary")
    dic("ColNameHeader") = bHdr
    Select Case charset
        Case "sjis"
            dic("CharacterSet") = 932
        Case "utf8"
            dic("CharacterSet") = 65001
        Case Else
    End Select
    Select Case UCase(dlm)
        Case "CSV", "TAB"
            dic("Format") = UCase(dlm) & "Delimited"
        Case Else
            dic("Format") = "Delimited(" & Left(dlm, 1) & ")"
    End Select
    mkIniHead = dicToStr(dic, "=", vbCrLf) & vbCrLf
End Function

Function mkIniTail(typeAry, Optional nameDic = Empty)
    If IsEmpty(nameDic) Then Set nameDic = CreateObject("Scripting.Dictionary")
    ret = ""
    Dim i
    Dim name
    i = 1
    For Each elm In typeAry
        name = IIf(nameDic.exists(i), nameDic(i), "F" & i)
        ret = ret & "Col" & i & "=" & name & " " & elm & vbCrLf
        i = i + 1
    Next elm
    mkIniTail = ret
End Function

Sub mkSchemaIniFile(fdr, fileAry, text)
    Dim stm
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stm = fso.CreateTextFile(fdr & "\Schema.ini")
    For Each fn In fileAry
        stm.WriteLine ("[" & fn & "]")
        stm.WriteLine (text)
    Next fn
    stm.Close
End Sub

Sub writeSchema(fdrn, fns, Optional defTbl As String = "schemaDef")
    Dim fso, stm
    Set fso = CreateObject("Scripting.FileSystemObject")
    schemaPath = fdrn & "\schema.ini"
    Set stm = fso.CreateTextFile(schemaPath)
    For Each fn In fns
        bn = Left(fn, InStr(fn, ".") - 1)
        stm.WriteLine ("[" & fn & "]")
        Call stm.WriteLine(TLookup(bn, defTbl, "def"))
    Next fn
    stm.Close
End Sub

Sub displayQueryTable(Optional sTbl = "", Optional sSQL = "", Optional spath = "", Optional top = "A4", Optional tblType = "")
    spath = getCurMdb(spath)
    Dim qt
    Dim sn, sCon, sSQL0
    Dim infoAry, r, c
    Sheets.Add After:=Sheets(Sheets.Count)
    sn = ActiveSheet.name
    sCon = "OLEDB;" & mdbConInfo(spath)
    sSQL0 = IIf(sSQL = "", "select * from " & sTbl, sSQL)
    Set qt = ThisWorkbook.Sheets(sn).QueryTables.Add(Connection:=sCon, Destination:=Sheets(sn).Range(top), sql:=sSQL0)
    qt.BackgroundQuery = False
    qt.Refresh
    qt.Delete
    Set qt = Nothing
    Sheets(sn).Range(top).CurrentRegion.Select
    Call Sheets(sn).ListObjects.Add(xlSrcRange, , , xlYes)
    '
    infoAry = mk2DAry(3, 2, "Path", spath, "SQL", sSQL0, "Type", tblType)
    r = Range(top).Row
    c = Range(top).Column
    Call lay2DAryAt(infoAry, r - 3, c)
End Sub

Sub displayODBCParamQuery(Optional sProc = "", Optional sSQL = "", Optional prmNameAry, Optional prmValAry, Optional prmTypeAry, Optional spath = "", Optional top = "A4")
    spath = getCurMdb
    Dim qt
    Dim sn, sCon, sSQL0
    Dim infoAry, r, c
    Sheets.Add After:=Sheets(Sheets.Count)
    sn = ActiveSheet.name
    sCon = "ODBC;" & mdbODBCInfo(spath)
    sSQL0 = IIf(sSQL = "", "select * from " & sTbl, sSQL)
    Set qt = ThisWorkbook.Sheets(sn).QueryTables.Add(Connection:=sCon, Destination:=Sheets(sn).Range(top))
    If sSQL <> "" Then
        qt.CommandType = xlCmdSql
        qt.CommandText = sql
    Else
        qt.CommandType = xlCmdSql
        qt.CommandText = TLookup(sProc, "procDef", "def")
    End If
    qt.BackgroundQuery = False
    prmNum1 = lenAry(prmNameAry)
    prmNum2 = lenAry(prmValAry)
    ReDim prms(0 To prmNum1 - 1)
    For i = 0 To prmNum1 - 1
        Set prms(i) = qt.Parameters.Add(prmNameAry(i)) ', prmTypeAry(i))
    Next i
    For i = 0 To prmNum2 - 1
        Call prms(i).SetParam(xlConstant, prmValAry(i))
    Next
    For i = prmNum2 To prmNum1 - 1
        Call prms(i).SetParam(xlPrompt, "enter " & prmNameAry(i))
    Next
    qt.Refresh
    qt.Delete
    'Set qt = Nothing
    Sheets(sn).Range(top).CurrentRegion.Select
    Call Sheets(sn).ListObjects.Add(xlSrcRange, , , xlYes)
    '
    infoAry = mk2DAry(3, 2, "Path", spath, "SQL", sSQL0, "Type", tblType)
    r = Range(top).Row
    c = Range(top).Column
    Call lay2DAryAt(infoAry, r - 3, c)
End Sub

Sub displayParamQuery(Optional sProc = "", Optional sSQL = "", Optional prmValAry, Optional spath = "", Optional top = "A4")
    Dim qt
    Dim sn, sCon, sSQL0
    Dim infoAry, r, c
    Dim cmd, rst
    Sheets.Add After:=Sheets(Sheets.Count)
    sn = ActiveSheet.name
    Call setCurMdb(spath)
    Call openMdbCon
    Set cmd = CreateObject("ADODB.Command")
    cmd.activeConnection = adoCon
    If sSQL = "" Then
        cmd.CommandText = sProc
        cmd.CommandType = 4 'adCmdStoredProc
    Else
        cmd.CommandText = sSQL
        cmd.CommandType = 1 'adCmdText
    End If
    Set rst = cmd.Execute(Parameters:=prmValAry)
    Set qt = ThisWorkbook.Sheets(sn).QueryTables.Add(Connection:=rst, Destination:=Sheets(sn).Range(top))
    qt.BackgroundQuery = False
    qt.Refresh
    qt.Delete
    Call closeCon
    Set cmd = Nothing
    'Set qt = Nothing
    Sheets(sn).Range(top).CurrentRegion.Select
    Call Sheets(sn).ListObjects.Add(xlSrcRange, , , xlYes)
    '
    infoAry = mk2DAry(3, 2, "Path", spath, "SQL", sSQL0, "Type", tblType)
    r = Range(top).Row
    c = Range(top).Column
    Call lay2DAryAt(infoAry, r - 3, c)
End Sub

Function mkInsertIntoSQL(tblTo, tblFrom, Optional colsTo = "", Optional colsFrom = "*", Optional where = "")
    Dim ret
    If colsTo <> "" Then colsTo = "(" & colsTo & ")"
    If where <> "" Then where = " Where " & where
    ret = "Insert Into " & tblTo & colsTo & " Select " & colsFrom & " From " & tblFrom & where
    mkInsertIntoSQL = ret
End Function

Function mkSelectIntoSQL(tblTo, tblFrom, Optional colsTo = "", Optional colsFrom = "*", Optional where = "")
    Dim ret
    If where <> "" Then where = " Where " & where
    ret = "Select " & mkSelectIntoCols(colsTo, colsFrom) & " Into " & tblTo & " From " & tblFrom & where
    mkSelectIntoSQL = ret
End Function

Function mkSelectIntoCols(colsTo, colsFrom)
    Dim ret
    Dim aryTo, aryFrom, aryTmp
    If colsTo = "" Then
        ret = colsFrom
    Else
        aryTo = Split(colsTo, ",")
        aryFrom = Split(colsFrom, ",")
        If UBound(aryTo) <> UBound(aryFrom) Then
            Call Err.Raise(1000, , "number of colsTo and colsFrom are different")
        End If
        aryTmp = eachJoinAry("", " as ", "", aryFrom, aryTo)
        ret = Join(aryTmp, ",")
    End If
    mkSelectIntoCols = ret
End Function

Function eachJoinAry(prefix, delm, suffix, ParamArray argArys())
    Dim arys
    Dim ret
    Dim arysL, arysD, aryL, aryD, aryIL
    arys = argArys
    arysL = LBound(arys)
    arysD = UBound(arys) - arysL
    aryL = LBound(arys(arysL))
    aryD = UBound(arys(arysL)) - aryL
    If aryD < 0 Then
        ret = Array()
    Else
        ReDim ret(0 To aryD)
        Dim i, j
        For j = 0 To aryD
            ret(j) = prefix
        Next j
        For i = 0 To arysD
            For j = 0 To aryD
                aryIL = LBound(arys(arysL + i))
                ret(j) = ret(j) & arys(arysL + i)(aryIL + j)
                ret(j) = ret(j) & IIf(i = arysD, suffix, delm)
            Next j
        Next i
    End If
    eachJoinAry = ret
End Function

Function eachConcateAry(aryC, ParamArray argArys())
    Dim ret
    Dim arys
    Dim arysL, arysD, aryL, aryD, aryIL
    arys = argArys
    arysL = LBound(arys)
    arysD = UBound(arys) - arysL
    aryL = LBound(arys(arysL))
    aryD = UBound(arys(arysL)) - aryL
    aryCL = LBound(aryC)
    aryCD = UBound(aryC) - aryCL
    If aryD < 0 Then
        ret = Array()
    Else
        ReDim ret(0 To aryD)
        Dim i, j
        For j = 0 To aryD
            ret(j) = aryC(aryCL)
        Next j
        For i = 0 To arysD
            For j = 0 To aryD
                aryIL = LBound(arys(arysL + i))
                ret(j) = ret(j) & arys(arysL + i)(aryIL + j)
                ret(j) = ret(j) & aryC(aryCL + i + 1)
            Next j
        Next i
    End If
    eachConcateAry = ret
End Function

Function dicToStr(dic, Optional dlmkey = ":", Optional dlmelm = ",") As String
    Dim ret As String
    ary = eachJoinAry("", dlmkey, "", dic.keys, dic.items)
    ret = Join(ary, dlmelm)
    dicToStr = ret
End Function

Function getXlTblName(xlTbl)
    Dim ret
    pos = InStr(xlTbl, "[")
    If pos = 0 Then
        ret = xlTbl
    Else
        ret = Left(xlTbl, pos - 1)
    End If
    getXlTblName = ret
End Function

Sub xlTblToCsv(xlTbl, Optional csvFdr = "", Optional bkn = "", Optional bPlusAbove = True, Optional bHdr = True, Optional intoAction = "select", Optional mdbPath = "")
    Dim tblFrom, tblTo, sql, csvPath
    Set fso = CreateObject("Scripting.FileSystemObject")
    If csvFdr = "" Then csvFdr = ThisWorkbook.path
    csvPath = csvFdr & "\" & getXlTblName(xlTbl) & ".csv"
    If intoAction = "select" Then
        If fso.fileexists(csvPath) Then fso.deletefile (csvPath)
    End If
    tblFrom = xlTblAsTable(xlTbl, bkn, bPlusAbove, bHdr)
    tblTo = txtAsTable(csvPath)
    Select Case LCase(intoAction)
        Case "select"
            sql = mkSelectIntoSQL(tblTo, tblFrom)
        Case "insert"
            sql = mkInsertIntoSQL(tblTo, tblFrom)
        Case Else
    End Select
    Set fso = CreateObject("Scripting.FileSystemObject")
    mdbPath = getCurMdb(mdbPath, , False)
    If Not fso.fileexists(mdbPath) Then mkMdb (mdbPath)
    Call execSQL(sql, mdbPath)
End Sub

Sub xlTblToMdb(xlTbl, Optional intoAction = "select", Optional mdbTbl = "", Optional bkn = "", Optional bPlusAbove = True, Optional bHdr = True, Optional mdbPath = "")
    Dim tblFrom, sql
    tblFrom = xlTblAsTable(xlTbl, bkn, bPlusAbove, bHdr)
    If bkn = "" Then bkn = ThisWorkbook.name
    If mdbTbl = "" Then mdbTbl = getXlTblName(xlTbl)
    Select Case LCase(intoAction)
        Case "select"
            sql = mkSelectIntoSQL(mdbTbl, tblFrom)
        Case "insert"
            sql = mkInsertIntoSQL(mdbTbl, tblFrom)
        Case Else
    End Select
    Call execSQL(sql, getCurMdb(mdbPath))
End Sub

Sub rangeToMdb(mdbTbl, xlPath, shn, rgn, Optional intoAction = "select", Optional mdbPath = "")
    Dim tblFrom, sql
    Set fso = CreateObject("Scripting.FileSystemObject")
    tblFrom = xlAsTable(xlPath, shn, rgn)
    Select Case LCase(intoAction)
        Case "select"
            sql = mkSelectIntoSQL(mdbTbl, tblFrom)
        Case "insert"
            sql = mkInsertIntoSQL(mdbTbl, tblFrom)
        Case Else
    End Select
    Call execSQL(sql, mdbPath)
End Sub

Function mkJoinSQL(tblA, tblB, colsSelectA, colsSelectB, colsJoinA, colsJoinB, Optional preJoin = "Inner", Optional where = "")
    Dim sql
    colsSA = Join(eachJoinAry("A.", "", "", Split(colsSelectA, ",")), ",")
    colsSB = Join(eachJoinAry("B.", "", "", Split(colsSelectB, ",")), ",")
    colsS = colsSA & IIf(colsSA <> "" And colsSB <> "", ",", "") & colsSB
    colsJ = eachConcateAry(Array("A.", " = B.", ""), Split(colsJoinA, ","), Split(colsJoinB, ","))
    If where <> "" Then where = " Where " & where
    sql = "Select " & colsS & " From " & tblA & " A " & preJoin & " Join " & tblB & " B On " & Join(colsJ, " and ") & where
    mkJoinSQL = sql
End Function

Function mkDiffSQL(tblA, tblB, colsSelectA, colsJoinA, colsJoinB)
    Dim sql
    sql = mkJoinSQL(tblA, tblB, colsSelectA, "", colsJoinA, colsJoinB, "Left", "B." & Split(colsJoinB, ",")(0) & " Is Null")
    mkDiffSQL = sql
End Function

Sub mkDiffView(mdbPath, tblA, tblB, colsSelectA, colsSelectB, colsJoinA, colsJoinB, Optional VA_B = "", Optional VB_A = "")
    If VA_B = "" Then VA_B = "Diff_" & tblA & "_" & tblB
    If VB_A = "" Then VB_A = "Diff_" & tblB & "_" & tblA
    Dim sqls(1 To 2)
    sqls(1) = "Create View " & VA_B & " As " & mkDiffSQL(tblA, tblB, colsSelectA, colsJoinA, colsJoinB)
    sqls(2) = "Create View " & VB_A & " As " & mkDiffSQL(tblB, tblA, colsSelectB, colsJoinB, colsJoinA)
    For Each sql In sqls
        Debug.Print sql
    Next
    Call execSQLs(sqls, mdbPath)
End Sub

Sub compactDB(spath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    'Set dbeg = CreateObject("DAO.DBEngine")
    ext = fso.GetExtensionName(spath)
    fdrn = fso.GetParentFolderName(spath)
    fn = fso.getfilename(spath)
    bn = fso.GetBaseName(spath)
    tpn = fdrn & "\" & bn & "_." & ext
    If fso.fileexists(tpn) Then fso.deletefile (tpn)
    Call DAO.DBEngine.CompactDatabase(spath, tpn)
    Call fso.CopyFile(tpn, spath)
    Call fso.deletefile(tpn)
    MsgBox "finished"
End Sub

Sub showDbCompactDialog()
    mdbPath = Application.GetOpenFilename("access file,*.mdb;*accdb", , "select file to be compacted")
    If TypeName(mdbPath) = "Boolean" Then Exit Sub
    Call compactDB(mdbPath)
End Sub

Sub mkView(sView, Optional mdbPath = "", Optional sXlTbl As String = "viewDef")
    sql0 = TLookup(sView, sXlTbl, "def")
    sql1 = addSpace(sql0)
    Call TSetUp(sql1, sView, sXlTbl, "def")
    sql2 = "Create View " & sView & " as " & vbLf & sql1
    Call getCurMdb(mdbPath)
    Call execSQL(sql2)
End Sub

Sub mkProc(sProc, Optional mdbPath = "", Optional sXlTbl As String = "procDef")
    Dim sql0, sql1, sql2, prm
    sql0 = TLookup(sProc, sXlTbl, "def")
    sql1 = addSpace(sql0)
    Call TSetUp(sql1, sProc, sXlTbl, "def")
    prm = Trim(TLookup(sProc, sXlTbl, "prm"))
    If prm = "" Then
        sql2 = "Create Procedure " & sProc & " as " & vbLf & sql1
    Else
        sql2 = "Create Procedure " & sProc & " as " & vbLf & " Parameters " & prm & ";" & vbLf & sql1
    End If
    Call execSQL(sql2, mdbPath)
End Sub

Function mkParamSQL(sProc, Optional sXlTbl As String = "procDef")
    Dim sql0, sql1, ret, prm
    sql0 = TLookup(sProc, sXlTbl, "def")
    sql1 = addSpace(sql0)
    prm = Trim(TLookup(sProc, sXlTbl, "prm"))
    If prm = "" Then
        ret = sql1
    Else
        ret = " Parameters " & prm & ";" & vbLf & sql1
    End If
    mkParamSQL = ret
End Function

Sub mkTable(sTable, Optional defName = "", Optional mdbPath = "", Optional sXlTbl As String = "tblDef")
    If defName = "" Then defName = sTable
    sql0 = TLookup(defName, sXlTbl, "def")
    sql1 = addComma(sql0)
    Call TSetUp(sql1, sTable, sXlTbl, "def")
    sql2 = "Create Table " & sTable & "(" & vbLf & sql1 & ")"
    Call getCurMdb(mdbPath)
    Call execSQL(sql2)
End Sub

Function addComma(txt, Optional rl = "r")
    Dim ret, ary0, ary1, i, j, ub
    ary0 = Split(txt, vbLf)
    ub = UBound(ary0)
    ReDim ary1(0 To ub)
    j = 0
    i = 0
    Do
        tmp = Trim(ary0(i))
        If tmp <> "" Then
            If Left(tmp, 1) = "," Then tmp = Right(tmp, Len(tmp) - 1)
            If Right(tmp, 1) = "," Then tmp = Left(tmp, Len(tmp) - 1)
            If tmp <> "" Then
                ary1(j) = tmp
                j = j + 1
            End If
        End If
        i = i + 1
    Loop Until i > ub
    If j = 0 Then
        ret = ""
    Else
        ReDim Preserve ary1(0 To j - 1)
        Select Case rl
            Case "r"
                ret = Join(ary1, vbLf & ",")
            Case "l"
                ret = Join(ary1, "," & vbLf)
            Case lese
        End Select
    End If
    addComma = ret
    ret = ""
End Function

Function addSpace(txt)
    Dim ret, ary0, i, ub
    ary0 = Split(txt, vbLf)
    ub = UBound(ary0)
    For i = 0 To ub
        ary0(i) = RTrim(ary0(i))
    Next i
    ret = Join(ary0, " " & vbLf)
    addSpace = ret
    ret = ""
End Function

Sub testas()
    txt = TLookup("p_rivalScoreView0", "procdef", "def")
    x = addSpace(txt)
    Debug.Print x
End Sub

Function getSqlVals(sql, Optional colNum = 1, Optional mdbPath = "")
    ReDim ret(0 To colNum - 1)
    Call getCurMdb(mdbPath)
    openMdbCon (getCurMdb(mdbPath))
    Set rst = CreateObject("adodb.recordset")
    Call rst.Open(sql, adoCon)
    If Not rst.EOF Then
        For i = 0 To colNum - 1
            ret(i) = rst(i)
        Next i
    End If
    rst.Close
    adoCon.Close
    getSqlVals = ret
End Function

Private Function lenAry(ary As Variant, Optional dm = 1) As Long
    lenAry = UBound(ary, dm) - LBound(ary, dm) + 1
End Function

Private Function TLookup(key, tbl As String, targetCol As String, Optional sourceCol As String = "", Optional otherwise = Empty, Optional bkn = "") As Variant
    Dim num, ret
    bkn0 = ActiveWorkbook.name
    If bkn = "" Then bkn = ThisWorkbook.name
    Workbooks(bkn).Activate
    Application.Volatile
    On Error GoTo lnError
    If sourceCol = "" Then sourceCol = Range(tbl & "[#headers]")(1, 1)
    num = WorksheetFunction.Match(key, Range(tbl & "[" & sourceCol & "]"), 0)
    If num = 0 Then
        ret = otherwise
    Else
        ret = Range(tbl & "[" & targetCol & "]")(num, 1)
    End If
    TLookup = ret
    Workbooks(bkn0).Activate
    Exit Function
lnError:
    Debug.Print Err.Description
    TLookup = Empty
    Workbooks(bkn).Activate
End Function

Private Sub TSetUp(vl, key, tbl As String, targetCol As String, Optional sourceCol As String = "", Optional bkn = "")
    bkn0 = ActiveWorkbook.name
    If bkn = "" Then bkn = ThisWorkbook.name
    Workbooks(bkn).Activate
    Application.Volatile
    On Error GoTo lnError
    If sourceCol = "" Then sourceCol = Range(tbl & "[#headers]")(1, 1)
    Range(tbl & "[" & targetCol & "]")(WorksheetFunction.Match(key, Range(tbl & "[" & sourceCol & "]"), 0), 1).Value = vl
    Workbooks(bkn0).Activate
    Exit Sub
lnError:
    Debug.Print Err.Description
End Sub

Private Function mk2DAry(r0, c0, ParamArray args())
    Dim ret, ary, i, r, c
    ary = args
    If lenAry(ary) <> r0 * c0 Then
        ret = Array()
    Else
        ReDim ret(0 To r0 - 1, 0 To c0 - 1)
        i = 0
        For Each elm In ary
            r = i \ c0
            c = i Mod c0
            ret(r, c) = elm
            i = i + 1
        Next elm
    End If
    mk2DAry = ret
End Function

Private Sub lay2DAryAt(ary, r, c)
    Dim cNum, rNum
    rNum = lenAry(ary)
    cNum = lenAry(ary, 2)
    Cells(r, c).Resize(rNum, cNum) = ary
End Sub

Sub exitForm(sFrm)
    Call Unload(UserForms(sFrm))
End Sub

Sub testss()
    x = getSqlVals("select count(*),avg(score) from scoreTbl", 2)
    printAry x
End Sub

Sub tetsa()
    x = mk2DAry(2, 3, 1, 2, 3, 4, 5, 6)
    printAry x
    Call lay2DAryAt(x, 2, 3)
End Sub
