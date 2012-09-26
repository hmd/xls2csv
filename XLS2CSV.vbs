option explicit

Class UploadFile
private m_MaxCol
private m_NowCol
private m_MaxRow
private m_NowRow
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
        m_NowCol = 1
        m_NowRow = 1
End Sub

public Function SetMaxCol(strCol)
        m_MaxCol = strCol
        ReDim Preserve m_TxtArray(m_MaxCol)
End Function

public Function SetMaxRow(strRow)
        m_MaxRow = strRow
End Function

public Function GetMaxRow()
        GetMaxRow = m_MaxRow
End Function

public Function GetMaxCol()
        GetMaxCol = m_MaxCol
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
        m_TxtArray(m_NowCol) = m_TxtStream
        if m_NowCol = m_MaxCol then
                uploadAdodbStream.WriteText(Join(m_TxtArray, ","))
                if m_NowRow < m_MaxRow then
                        uploadAdodbStream.WriteText(vbcrlf)
                end if
                m_NowCol = 1
                m_NowRow = m_NowRow + 1
        else
                m_NowCol = m_NowCol + 1
        end if
end Function

Public sub Save(strFileName)
        uploadAdodbStream.SaveToFile strFileName, adSaveCreateOverWrite ' force overwrite
        uploadAdodbStream.Close()
end sub

Private Sub Class_Terminate
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
dim gStartLine
dim gChkDigit
dim chkError
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
        gStartLine = args.Item(2)
        gChkDigit = args.Item(4)
End If
call ChkArgs(chkError)
If chkError then
        WScript.Quit 10
else
        Call CrtData()
end if
set objRE = nothing
set objStr = nothing
WScript.Quit 10
'------------------------------------------------------------------------------------------------------
sub ChkArgs(byref chkError)
        dim FolderName
        If FSO.FileExists(gXlsPath) = False Then
                MsgBox gXlsPath & " is nothing", vbCritical
                chkError = True
        else
                if Chkfiletype(gXlsPath,"xls")  then chkError = true
        End If

        If FSO.FileExists(gFdfPath) = False Then
                MsgBox gFdfPath & " is nothing", vbCritical
                chkError = True
        else
                if Chkfiletype(gFdfPath,"fdf") then chkError = true
        End If

        If gDataPath = Empty Then 
                xlsName = FSO.getfilename(gXlsPath)
                FolderName = Replace(gXlsPath, xlsName, "")
                gDataPath = FolderName & Replace(xlsName, ".xls", ".csv")
        else
                FolderName = Replace(gDataPath, FSO.getfilename(gDataPath), "")
                If FSO.FolderExists(FolderName) = False Then
                        MsgBox FolderName & "is nothing",vbcritical
                        chkError = True
                End If
        end if
        if gChkDigit <> "T" and gChkDigit <> "F" then
                MsgBox "Parameter 5th is incorrect." & vbcrlf & "Correct value is  'T or F'",vbcritical
                chkError = True
        end if

        If FSO.FileExists(gDataPath) Then FSO.DeleteFile (gDataPath)
end sub
                        '------------------------------------------------------------------------------------------------------
                        sub CrtData()
                                Dim xSh
                                Dim t,c
                                Dim oBxl : Set oBxl = CreateObject("Excel.Application")
                                Dim uploadText : Set uploadText = new UploadFile
                                Dim wrkResult
                                'Get FDF info
                                call GetFdfinfo(gFdfPath,gType,gLength,gDecimal)
                                'Get xls data
                                oBxl.Visible = False
                                oBxl.Workbooks.Open gXlsPath
                                set xSh = oBxl.ActiveWorkbook.ActiveSheet
                                With xSh.UsedRange
                                        uploadText.SetMaxRow(.Rows(.Rows.Count).Row)
                                end with
                                uploadText.SetMaxCol(UBound(gType))
                                'oBxl.Quit
                                'Check data row
                                On Error Resume Next
                                if uploadText.GetMaxRow < cint(gStartLine)   Then
                                        If Err.Number <> 0 Then
                                                MsgBox (gXlsPath & " is  no data"),vbCritical
                                        else 
                                                MsgBox(gXlsPath & " is no data" & vbcr & "after " & gStartLine & " lines"),vbCritical
                                        end if
                                        On Error Goto 0
                                Else
                                        On Error Goto 0

                                        'Start exchanging
                                        For t = gStartLine To uploadText.GetMaxRow
                                                For c = 1 To UBound(gType)
                                                        On Error Resume Next
                                                        if ChkData(xSh.Cells(t, c), gType(c),gLength(c),gDecimal(c),gChkDigit,t,c) then
                                                                If Err.Number <> 0 Then
                                                                        MsgBox t & "row " & c & "column" & "  has Error data", vbCritical
                                                                        tf.Close
                                                                        FSO.DeleteFile gDataPath
                                                                        WScript.Quit 10
                                                                end if
                                                                tf.Close
                                                                FSO.DeleteFile gDataPath
                                                                WScript.Quit 10
                                                        else
                                                                'VBScriptは引数が2つ以上あるかつ、引数に括弧が使用されている場合
                                                                '戻り値を設定しないとエラーになる仕様
                                                                wrkReslt = uploadText.SetData(xSh.Cells(t, c), gType(c))
                                                                uploadText.WriteText()
                                                        End If
                                                        On Error Goto 0
                                                Next
                                        Next 
                                        uploadText.Save(gDataPath)
                                        'If Err.Number <> 0 Then
                                        '        msgbox "Failed to save file", vbcritical
                                        'End If
                                End If
                        end sub
                        '------------------------------------------------------------------------------------------------------
                        'Check data type function
                        Function ChkData(xlsString,byval sType,byval lenNumber,byval lenDecimal,gChkDigit,Row,clomun)
                                if InStr( xlsString, vbCrLf ) or _
                                        InStr( xlsString, vbCr ) or _   
                                        InStr( xlsString, vbLf )  then
                                        MsgBox Row & "row " & clomun & "column" & "  has line break", vbCritical
                                        ChkData = True
                                        exit Function
                                elseif sType = "1" Then  'Character
                                        if CalcByte(xlsString) > lenNumber and gChkDigit = "T" then
                                                MsgBox Row & "row " & clomun & "column" & "  is overflow", vbCritical
                                                ChkData = True
                                                exit Function
                                        end if
                                elseif sType = "2" Then 'Numeric
                                        If xlsString <> "" then 
                                                if  vartype(xlsString) > 5 Then
                                                        MsgBox Row & "row " & clomun & "column" & "  is not numeric", vbCritical
                                                        ChkData = True
                                                        Exit Function
                                                elseif gChkDigit = "T" then 
                                                        strKeta = split(xlsString,".")
                                                        if lenDecimal > 0 then
                                                                lenNumber = lenNumber - (2 + lenDecimal)
                                                        else
                                                                lenNumber = lenNumber - 1
                                                        end if
                                                        if ubound(strKeta) = 1 then 
                                                                if CalcByte(strketa(0)) > lenNumber or _
                                                                        CalcByte(strketa(1)) > lenDecimal then
                                                                        MsgBox Row & "row " & clomun & "column" & "  is overflow", vbCritical
                                                                        ChkData = True
                                                                        exit Function
                                                                end if
                                                        else
                                                                if CalcByte(strketa(0)) > lenNumber then
                                                                        MsgBox Row & "row " & clomun & "column" & "  is overflow", vbCritical
                                                                        ChkData = True
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
                                                        if slashPosition > 0 then
                                                                shousu = Mid(strString(3), slashPosition + 1)
                                                                seisu = Mid(strString(3),1, slashPosition - 1)
                                                        else
                                                                shousu = 0
                                                                seisu = strString(3)
                                                        end if

                                                        gType(i) = cint(strString(2)) 'data type
                                                        gLength(i) = cint(seisu) 'Length
                                                        gDecimal(i) = cint(shousu) 'Length
                                                end if
                                        else
                                                'Client Acess
                                                objRE.Pattern = "^Length*"
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
                        Function CalcByte(ByVal a)
                                Dim c
                                c = 0
                                Dim i
                                For i = 0 To Len(a) - 1
                                        Dim k
                                        k = Mid(a, i + 1, 1)
                                        If (Asc(k) And &HFF00) = 0 Then
                                                c = c + 1
                                        Else
                                                c = c + 2
                                        End If
                                Next
                                CalcByte = c
                        End Function
