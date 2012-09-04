option explicit
Dim gType()
Dim gLength()
Dim gDecimal()
Dim args
dim gXlsPath 
dim gFdfPath 
dim gDataPath
dim gDataLine
dim gChkDigit
dim chkError
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
        gDataLine = args.Item(2)
        gChkDigit = args.Item(4)
End If
call ChkArgs(chkError)
If chkError then
        WScript.Quit 10
else
        Call CrtData()
end if
WScript.Quit 10
'------------------------------------------------------------------------------------------------------
sub ChkArgs(byref chkError)

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
                dataName = FSO.getfilename(gDataPath)
                FolderName = Replace(gDataPath, dataName, "")
                If FSO.FolderExists(FolderName) = False Then
                        MsgBox FolderName & "is nothing",vbcritical
                        chkError = True
                End If
        end if
        if gChkDigit <> "Y" and gChkDigit <> "F" then
                MsgBox "Parameter 5th is incorrect." & vbcrlf & "Correct value is  'Y or F'",vbcritical
                chkError = True
        end if

If FSO.FileExists(gDataPath) Then FSO.DeleteFile (gDataPath)
end sub
'------------------------------------------------------------------------------------------------------
sub CrtData()

        Dim oBxl : Set oBxl = CreateObject("Excel.Application")
        'Get FDF info
        call GetFdfinfo(gFdfPath,gType,gLength,gDecimal)
        'Get xls data
        oBxl.Visible = False
        oBxl.Workbooks.Open gXlsPath
        set xSh = oBxl.ActiveWorkbook.ActiveSheet
        xlsEndRow = xSh.UsedRange 
        xlsData = xSh.Range(xSh.Cells(1,1),xSh.Cells(ubound(xlsEndRow),ubound(gType))).value 
        oBxl.Quit
        Erase xlsEndRow
        'Check data row
        On Error Resume Next
        if ubound(xlsData) < cint(gDataLine)   Then
                If Err.Number <> 0 Then
                        MsgBox (gXlsPath & " is  no data"),vbCritical
                else 
                        MsgBox(gXlsPath & " is no data" & vbcr & "after " & gDataLine & " lines"),vbCritical
                end if
                On Error Goto 0
        Else
                On Error Goto 0

                'Start exchanging
                For t = gDataLine To UBound(xlsData)
                        For c = 1 To UBound(gType)
                                On Error Resume Next
                                if ChkData(xlsData(t, c), gType(c),gLength(c),gDecimal(c),gChkDigit,t,c) then
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
                                        call edtData(xlsData(t, c), gType(c)) 
                                End If
                                On Error Goto 0
                        Next
                Next 
                call savefile(gDataPath,xlsdata,gDataLine,ubound(gType))
                If Err.Number <> 0 Then
                        msgbox "Failed to save file", vbcritical
                End If
        End If
end sub
'------------------------------------------------------------------------------------------------------
'Convert data function
Function edtData(xlsString,sType)
        If sType = "1" Then 
                xlsString = replace(xlsString,"""","""""")
                xlsString = """" & xlsString & """"
        elseIf sType = "2" Then
                If xlsString = "" Then
                        xlsString = 0
                Else
                        xlsString = xlsString
                End If
        End If
End Function
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
                if CalcByte(xlsString) > lenNumber and gChkDigit = "Y" then
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
                        elseif gChkDigit = "Y" then 
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
        Set objInFile = FSO.OpenTextFile(gFdfPath, 1)
        Set objRE = CreateObject("VBScript.RegExp")
        set objStr =  CreateObject("ADODB.Stream")
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
        set objRE = nothing
        set objStr = nothing
end Function
'-------------------------------------------------------------------------------------------------------------------
function saveFile(filename, text,strRow,endColoum)
        On Error Resume Next

        ' ADODB.Streamのモード
        Dim adTypeBinary : adTypeBinary = 1
        Dim adTypeText : adTypeText = 2
        Dim adSaveCreateOverWrite : adSaveCreateOverWrite = 2

        ' ADODB.Streamを作成
        Dim pre : Set pre = CreateObject("ADODB.Stream")
        ' 最初はテキストモードでUTF-8で書き込む
        pre.Type = adTypeText
        pre.Charset = "shift-jis"
        pre.Open()
        for i = strRow to ubound(text) 
                for j = 1 to endColoum
                        pre.WriteText(text(i,j))
                        if j < endColoum then pre.WriteText(",")
                        if j = endColoum and i < ubound(text) then pre.WriteText(vbcrlf)
                        next
                next
                pre.SaveToFile filename, adSaveCreateOverWrite ' force overwrite
                pre.Close()
        End function
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
