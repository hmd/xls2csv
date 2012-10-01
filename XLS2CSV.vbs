option explicit

Class UploadFile
private m_MaxCol
private m_CurrentCol
private m_MaxRow
private m_CurrentRow
private m_TxtStream
private adTypeBinary
private adTypeText
private adSaveCreateOverWrite
private uploadAdodbStream
private m_RegExp
private m_TxtArray()

Private Sub Class_Initialize
        ' ADODB.Streamのモード
        set uploadAdodbStream = CreateObject("ADODB.Stream")
        Set m_RegExp = CreateObject("VBScript.RegExp")
        m_RegExp.Pattern = """"
        adTypeBinary = 1
        adTypeText = 2
        adSaveCreateOverWrite = 2
        uploadAdodbStream.Type = adTypeText
        uploadAdodbStream.Charset = "shift-jis"
        uploadAdodbStream.Open()
        m_CurrentCol = 0
        m_CurrentRow = 1
End Sub

public Function SetMaxCol(strCol)
        m_MaxCol = strCol - 1
        ReDim Preserve m_TxtArray(m_MaxCol)
End Function

public Function SetMaxRow(strRow)
        m_MaxRow = strRow
End Function

Public Function SetData(strData,strType) 
        If strType = "1" Then 
                m_TxtStream = """" & m_RegExp.Replace(strData,"""""") & """"
        elseIf strType = "2" Then
                If strData = "" Then
                        m_TxtStream = 0
                Else
                        m_TxtStream = strData
                End If
        End If
end Function

Public Function WriteText() 
        m_TxtArray(m_CurrentCol) = m_TxtStream
        if m_CurrentCol = m_MaxCol then
                uploadAdodbStream.WriteText(Join(m_TxtArray, ","))
                if m_CurrentRow < m_MaxRow then
                        uploadAdodbStream.WriteText(vbcrlf)
                end if
                m_CurrentCol = 0
                m_CurrentRow = m_CurrentRow + 1
        else
                m_CurrentCol = m_CurrentCol + 1
        end if
end Function

Public sub Save(strFileName)
        uploadAdodbStream.SaveToFile strFileName, adSaveCreateOverWrite ' force overwrite
        uploadAdodbStream.Close()
end sub

Private Sub Class_Terminate
        set uploadAdodbStream = nothing
        Set m_RegExp = nothing
End Sub

end Class

'----------------------------------------------------------------------------------------------------
Dim gType()
Dim gLength()
Dim gDecimal()
Dim args
dim gXlsPath 
dim gFdfPath 
dim gDataPath
dim gStartRow
dim gChkLength
dim ArgErrExists
Dim objRE : Set objRE = CreateObject("VBScript.RegExp")
Dim objStr : set objStr =  CreateObject("ADODB.Stream")
Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments
'
If args.Count <> 5 Then
        MsgBox ("Please set parameter correct")
        WScript.Quit 10
Else
        gXlsPath = args.Item(0)
        gFdfPath = args.Item(1)
        gDataPath = args.Item(3)
        gStartRow = args.Item(2)
        gChkLength = args.Item(4)
End If
call ChkArgs(ArgErrExists)
If ArgErrExists = False then
        Call CreateCsv()
end if
set objRE = nothing
set objStr = nothing
set FSO = nothing
WScript.Quit 10
'------------------------------------------------------------------------------------------------------
sub ChkArgs(byref ArgErrExists)
        dim FolderName
        If FSO.FileExists(gXlsPath) Then
                if Chkfiletype(gXlsPath,"xls")  then ArgErrExists = true
        else
                MsgBox gXlsPath & " is nothing", vbCritical
                ArgErrExists = True
        End If

        If FSO.FileExists(gFdfPath) Then
                if Chkfiletype(gFdfPath,"fdf") then ArgErrExists = true
        else
                MsgBox gFdfPath & " is nothing", vbCritical
                ArgErrExists = True
        End If

        If gDataPath = Empty Then 
                xlsName = FSO.getfilename(gXlsPath)
                FolderName = Replace(gXlsPath, xlsName, "")
                gDataPath = FolderName & Replace(xlsName, ".xls", ".csv")
        else
                FolderName = Replace(gDataPath, FSO.getfilename(gDataPath), "")
                If FSO.FolderExists(FolderName) = False Then
                        MsgBox FolderName & "is nothing",vbcritical
                        ArgErrExists = True
                End If
        end if
        if gChkLength <> "T" and gChkLength <> "F" then
                MsgBox "Parameter 5th is incorrect." & vbcrlf & "Correct value is  'T or F'",vbcritical
                ArgErrExists = True
        end if

        If FSO.FileExists(gDataPath) Then FSO.DeleteFile (gDataPath)
end sub
'------------------------------------------------------------------------------------------------------
sub CreateCsv()
        Dim xSh
        Dim t,c
        Dim uploadText : Set uploadText = new UploadFile
        Dim wrkResult
        Dim xlsObject
        Dim xlsEndRow,xlsEndCol
        Dim StrXlsStream
        'Get FDF info
        call GetFdfinfo(gFdfPath,gType,gLength,gDecimal)
        Set xlsObject = WScript.GetObject(gXlsPath, "Excel.Sheet")
        xlsObject.Application.ScreenUpdating = False
        'Get xls data
        set xSh = xlsObject.Sheets(1)
        With xSh.UsedRange
                xlsEndRow = .Rows(.Rows.Count).Row
                uploadText.SetMaxRow(xlsEndRow)
        end with
        xlsEndCol = UBound(gType)
        uploadText.SetMaxCol(xlsEndCol)
        'Check data row
        On Error Resume Next
        if xlsEndRow  < cint(gStartRow)   Then
                If Err.Number <> 0 Then
                        MsgBox (gXlsPath & " is  no data"),vbCritical
                else 
                        MsgBox(gXlsPath & " is no data" & vbcr & "after " & gStartRow & " lines"),vbCritical
                end if
        Else
                'Start exchanging
                For t = gStartRow To xlsEndRow 
                        StrXlsStream = xSh.Range(xSh.Cells(t,1),xSh.Cells(t,xlsEndCol)).value
                        For c = 1 To xlsEndCol
                                if ChkXlsData(StrXlsStream(1,c), gType(c),gLength(c),gDecimal(c),gChkLength,t,c) then
                                        If Err.Number <> 0 Then
                                                MsgBox t & "row " & c & "column" & "  has Error data", vbCritical
                                        end if
                                        xlsObject.Application.ScreenUpdating = True
                                        xlsObject.Close
                                        set xlsObject = nothing
                                        set uploadText = nothing
                                        Exit sub
                                else
                                        'VBScriptは引数が2つ以上あるかつ、引数に括弧が使用されている場合
                                        '戻り値を設定しないとエラーになる仕様
                                        wrkResult = uploadText.SetData(StrXlsStream(1,c), gType(c))
                                        uploadText.WriteText()
                                End If
                        Next
                Next 
                uploadText.Save(gDataPath)
        End If
        xlsObject.Application.ScreenUpdating = True
        xlsObject.Close
        set xlsObject = nothing
        set uploadText = nothing
        On Error Goto 0
end sub
'------------------------------------------------------------------------------------------------------
'Check data type function
Function ChkXlsData(xlsString,byval sType,byval lenNumber,byval lenDecimal,gChkLength,Row,clomun)
        Dim strKeta
        objRE.Pattern = "(\n|\r)"
        if objRE.test(xlsString) then
                MsgBox Row & "row " & clomun & "column" & "  has line break", vbCritical
                ChkXlsData = True
                exit Function
        elseif sType = "1" Then  'Character
                if CalcByte(xlsString) > lenNumber and gChkLength = "T" then
                        MsgBox Row & "row " & clomun & "column" & "  is overflow", vbCritical
                        ChkXlsData = True
                        exit Function
                end if
        elseif sType = "2" Then 'Numeric
                If xlsString <> "" then 
                        if  vartype(xlsString) > 5 Then
                                MsgBox Row & "row " & clomun & "column" & "  is not numeric", vbCritical
                                ChkXlsData = True
                                Exit Function
                        elseif gChkLength = "T" then 
                                strKeta = split(Cstr(xlsString),".")
                                if ubound(strKeta) = 1 then 
                                        if CalcByte(strketa(0)) > lenNumber or _
                                                CalcByte(strketa(1)) > lenDecimal then
                                                MsgBox Row & "row " & clomun & "column" & "  is overflow", vbCritical
                                                ChkXlsData = True
                                                exit Function
                                        end if
                                else
                                        if CalcByte(strketa(0)) > lenNumber then
                                                MsgBox Row & "row " & clomun & "column" & "  is overflow", vbCritical
                                                ChkXlsData = True
                                                exit Function
                                        end if
                                end if
                        End If
                End If
        End If
End Function
'------------------------------------------------------------------------------------------------------
'Check file type
Function Chkfiletype(Path,fileType) 
        If lcase(FSO.GetExtensionName(Path)) <> fileType Then
                MsgBox Path & " is not " & fileType & " file",vbcritical
                Chkfiletype = True
        End If
end Function
'------------------------------------------------------------------------------------------------------
'Get fdf information
Function GetFdfinfo(gFdfPath,byref gType,byref gLength,byref gDecimal) 
        Dim objInFile : Set objInFile = FSO.OpenTextFile(gFdfPath, 1)
        Dim strString
        Dim fdfAry
        Dim slashPosition
        Dim shousu
        Dim seisu
        Dim j,i
        With objStr
                .Open
                .Charset = "_autodetect_all"
                .LoadFromFile(gFdfPath)
                fdfAry = split(.ReadText,vbcrlf)
                .close
        End With
        for j =3 to ubound(fdfAry)
                strString = Split(fdfAry(j), " ")
                objRE.Pattern = "^P"
                if objRE.Test(fdfAry(j)) then 
                        if  UBound(strString) = 3 then
                                'P-Comm
                                i = i + 1
                                ReDim Preserve gType(i)
                                ReDim Preserve gLength(i)
                                ReDim Preserve gDecimal(i)
                                slashPosition = InStrRev(strString(3), "/")
                                if strString(2) = 1 then 'String type
                                        shousu = 0
                                        seisu = strString(3)
                                else 'Numeric type
                                        if slashPosition > 0 then
                                                shousu = Mid(strString(3), slashPosition + 1)
                                                seisu = Mid(strString(3),1, slashPosition - 1) - (shousu + 2)
                                        else
                                                shousu = 0
                                                seisu = strString(3) - 1
                                        end if
                                end if
                                gType(i) = cint(strString(2)) 'data type
                                gLength(i) = cint(seisu) 'Length
                                gDecimal(i) = cint(shousu) 'Length
                        end if
                else
                        'Client Acess
                        objRE.Pattern = "^Length=*"
                        if objRE.Test(fdfAry(j)) then
                                i = i + 1
                                ReDim Preserve gType(i)
                                ReDim Preserve gLength(i)
                                ReDim Preserve gDecimal(i)
                                gLength(i) = cint(replace(fdfAry(j),"Length=",""))
                        else
                                objRE.Pattern = "^Scale=*"
                                if objRE.Test(fdfAry(j)) then
                                        gDecimal(i) = cint(replace(fdfAry(j),"Scale=","")) '
                                        gLength(i) = gLength(i) - (2 + gDecimal(i))
                                else
                                        objRE.Pattern = "^Type=*"
                                        if objRE.Test(fdfAry(j)) then
                                                gType(i) = cint(replace(fdfAry(j),"Type=","")) '
                                        end if
                                end if
                        end if
                end if
        next
        objInFile.Close
        set objInfile = nothing
end Function
'-------------------------------------------------------------------------------------------------------------------
Function CalcByte(TargetStream)
        Dim i, cnt
        Dim StringType
        Dim DBCSSV
        cnt = 0
        For i = 1 To Len(TargetStream)
                StringType = Asc(Mid(TargetStream, i, 1))
                If (0 < StringType And StringType < 255) Then
                        '0f　漢字の終了
                        If (DBCSSV = 1) Then
                                cnt = cnt + 1
                        End If
                        cnt = cnt + 1
                        DBCSSV = 0

                Else
                        '0e 漢字の開始
                        If (DBCSSV = 0) Then
                                cnt = cnt + 1
                        End If
                        DBCSSV = 1
                        cnt = cnt + 2
                End If
        Next
        If (DBCSSV >= 1) Then
                cnt = cnt + 1
        End If
        CalcByte = cnt
End Function
