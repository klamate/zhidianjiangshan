Set oArgs = WScript.Arguments
Dim fso

Function DealXLSX2CSV(fileDir, fileXlsx, fileCsv, strSheet, iDeleteLine)
  if InStr(1, fileCsv, ":", 1) > 0 then
    strSaveToFile = fileCsv
  else
    strSaveToFile = fileDir & fileCsv
  end if
    
  If fso.fileExists(strSaveToFile) Then
    fso.DeleteFile(strSaveToFile)
  End If
  
  ' MsgBox( fileDir & " " & fileXlsx & " To " & fileCsv )
  set objExcel = createobject("excel.application")
  set objworkbook = objExcel.workbooks.open( fileDir & fileXlsx )
  
  if Len(strSheet) > 0 then
    objworkbook.Worksheets(strSheet).Activate
  end if
  
  set SSetmp = objworkbook.ActiveSheet
  with SSetmp
      .UsedRange.Copy 
      .UsedRange.PasteSpecial -4163
      .Rows("1:" & CStr(iDeleteLine) ).Delete
  end with
    
  call objExcel.activeworkbook.saveas( strSaveToFile, 6)
  objworkbook.Saved = true
  objworkbook.close
  objExcel.quit
  
  If fso.fileExists(strSaveToFile) Then
    set ws = WScript.CreateObject("WScript.Shell")
    strToolPath = Wscript.ScriptFullName
    strFileExt = Split( strToolPath, "\", -1, 1 )
    iArrayCount = UBound(strFileExt) - LBound(strFileExt)
    strFileExt(iArrayCount) = ""
    strToolPath = ""
    For Each s In strFileExt
      strToolPath = strToolPath & s & "\"
    Next
    strToolPath = Left( strToolPath, Len(strToolPath)-1 )
    strTargetPath = strSaveToFile
    strToolFull = """" & strToolPath & "ConverAnsiToUTF8.bat""" & " " & """" & strTargetPath & """" & " " & """" & strTargetPath & """"
    ws.Run strToolFull,0,false
  end if
End Function

If oArgs.count = 0 Then
  MsgBox("ÇëÊäÈë²ÎÊý")
Else
  Set fso = CreateObject("Scripting.FileSystemObject")
  Dim strPath, iDeleteLine, strFileExt
  strPath = oArgs(0)
  
  strSheetCaption = ""
  If oArgs.count >= 2 then
    strSheetCaption = oArgs(1)
  End IF
  
  If oArgs.count >= 4 then
    iDeleteLine = CInt(oArgs(3))
  else
    iDeleteLine = 1
  End If
  
  If fso.folderExists(strPath) Then
  Else
    If fso.fileExists(strPath) Then
      strFileExt = Split( strPath, ".", -1, 1 )
      iArrayCount = UBound(strFileExt) - LBound(strFileExt)
      if UCase(strFileExt(iArrayCount)) = UCase("xlsx") then
        strFileExt = replace(strPath, "\", "/")
        strFileExt = Split( strFileExt, "/", -1, 1 )
        iArrayCount = UBound(strFileExt) - LBound(strFileExt)
        strFileFullname = strFileExt(iArrayCount)
        strFileExt(iArrayCount) = ""
        strFileDir = ""
        For Each s In strFileExt
          strFileDir = strFileDir & s & "\"
        Next
        strFileDir = Left( strFileDir, Len(strFileDir)-1 )
        
        strFileExt = Split( strFileFullname, ".", -1, 1 )
        iArrayCount = UBound(strFileExt) - LBound(strFileExt)
        strFileExt(iArrayCount) = "csv"

        strFileRename = ""
        For Each s In strFileExt
          strFileRename = strFileRename & s & "."
        Next
        
        strFileRename = Left( strFileRename, Len(strFileRename)-1 )
        If oArgs.count >= 3 then
          strFileRename = oArgs(2)
        End If
          
        DealXLSX2CSV strFileDir, strFileFullname, strFileRename, strSheetCaption, iDeleteLine
      end if
    End If
  End If
End If

'For Each s In oArgs
'    MsgBox(s)
'Next
    
Set oArgs = Nothing