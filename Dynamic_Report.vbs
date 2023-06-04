Option Explicit

Call Main()
Sub Main()
    Dim oEdge, oShell, by, elemSelect, oFso, oFile, oFolder, sFolderPath, sFilePath, sDownloadFile
    Dim sLinkList, sLink, sTypeName, nowTime, sStartDate, sEndDate, iloop, sUserProrfile, LoopCnt
    
    Set oShell = CreateObject("WScript.Shell")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Set oEdge = CreateObject("Selenium.WebDriver")
    Set by = CreateObject("Selenium.by")
    sUserProrfile = oShell.ExpandEnvironmentStrings("%userprofile%")
    sFolderPath = sUserProrfile & "\Downloads\"
    sFilePath = sFolderPath & "Unconfirmed*"
    Set oFolder = oFso.GetFolder(sFolderPath)
    nowTime = Now() - TimeSerial(10, 30, 0)
    
    For Each oFile In oFolder.Files
        If InStr(oFile.Name, "Dynamic") > 0 Then
            oFso.DeleteFile oFile
        End If
    Next
    
    sStartDate = Month(nowTime) & "/" & Day(nowTime) & "/" & Year(nowTime) & " 00:00"
    sEndDate = Month(nowTime) & "/" & Day(nowTime) & "/" & Year(nowTime) & " 23:59"
    oEdge.Start "Edge", "https://omega2010.uhc.com/"
    oEdge.Get "https://omega2010.uhc.com/"
    If oEdge.IsElementPresent(by.Css("#login"), 3) Then oEdge.FindElement(by.Css("#login")).Click
    oEdge.Get "https://omega2010.uhc.com/Administrator/Report/DynamicProcessedWorkItem"
    oEdge.wait 500
    Set elemSelect = oEdge.FindElementByCss("#BusinessArea").AsSelect
    elemSelect.SelectByText "CSS UNET"
    oEdge.FindElementByCss("#SearchStartDateTime").Clear
    oEdge.FindElementByCss("#SearchStartDateTime").SendKeys sStartDate
    oEdge.FindElementByCss("#SearchEndDateTime").Clear
    oEdge.FindElementByCss("#SearchEndDateTime").SendKeys sEndDate
    Set elemSelect = oEdge.FindElementByCss("#TimeZone").AsSelect
    elemSelect.SelectByText "Central Standard Time"
    
    For iloop = 1 To 3
        Set elemSelect = oEdge.FindElementByCss("#SelectedWorkTypeId").AsSelect
        If iloop = 1 Then elemSelect.SelectByText "CSS CRT": sTypeName = "CSS CRT"
        If iloop = 2 Then elemSelect.SelectByText "CSS UNET Review": sTypeName = "CSS UNET Review"
        If iloop = 3 Then elemSelect.SelectByText "CSS UNET Rework": sTypeName = "CSS UNET Rework"
        sTypeName = Right(sTypeName, Len(sTypeName) - InStrRev(sTypeName, " "))
        oEdge.wait 500
        oEdge.FindElementByCss("#cloneBtn").Click
        oEdge.wait 500
        oEdge.FindElementByCss("#btnSubmit").Click
        oEdge.wait 10000
        oEdge.SwitchToFrame "riframe"
        oEdge.wait 1000
        Do While oEdge.FindElementByCss("#ctl00_cphContent_rvOmegaSumaryReports_ctl09").IsPresent = True
            oEdge.wait 1000
            If InStr(oEdge.FindElementByCss("#ctl00_cphContent_rvOmegaSumaryReports_ctl09").Text, "DYNAMIC PROCESSED WORK ITEM REPORT") > 0 Then
                Exit Do
            End If
            DoEvents
        Loop
        oEdge.wait 1000
        Set sLinkList = oEdge.FindElementsByTag("a")
        For Each sLink In sLinkList
             If sLink.Attribute("title") = "Excel" Then
                oEdge.ExecuteScript "arguments[0].click();", sLink
                Exit For
            End If
        Next
        oEdge.SwitchToParentFrame
        oEdge.wait 2000
        
        LoopCnt = 0: sDownloadFile = ""
        Do While True
            For Each oFile In oFolder.Files
                If instr(oFile.Name,".crdownload") > 0 Then
                    sDownloadFile = oFile.Path
                        Exit For
                End If
            Next
            If sDownloadFile = "" Then Exit Do
            sDownloadFile = ""
            Sleep 1000
            LoopCnt = LoopCnt + 1
            If LoopCnt > 30 Then MsgBox "Omega file download terminated unexpectedly": Exit Sub
        Loop
                
        For Each oFile In oFolder.Files
            If InStr(Left(oFile.Name, 7), "Dynamic") > 0 Then
                oFso.GetFile(oFile).Name = sTypeName & " " & oFile.Name
                    Exit For
            End If
        Next
    Next
    oEdge.Quit
    
    Call ShareEmail(sFolderPath)
End Sub


Private Sub ShareEmail(folder_path)
    Dim outlook, email, sText1, sText2, sCopybody, folder, file, fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folder_path)
    Set outlook = CreateObject("Outlook.Application")
    Set email = outlook.CreateItem(0)
    sText1 = "<font face=""Calibri""><br>Hi All,<br><br>Here by attached Omega Dynamic Production Report<br></font>"
    sText2 = "<font face = ""Calibri"">Regards,</font>"
    email.Display
    sCopybody = email.HTMLBody
    email.To = "UnetSME_DL@ds.uhc.com"
    email.CC = "UBH_AM_S_DL@ds.uhc.com"
    email.Subject = "Omega Dynamic Report " & FormatDateTime(Now, vbLongTime)
    email.HTMLBody = sText1 & " <br> " & sText2 & " <br> " & sCopybody
    For Each file In folder.Files
        If instr(file.Name, "Dynamic") > 0 Then
            email.Attachments.Add folder_path & "\" & file.Name
        End If
    Next
    email.Display
    Set email = Nothing
    Set outlook = Nothing
End Sub
