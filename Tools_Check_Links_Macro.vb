' Found here: https://stackoverflow.com/questions/22256522/checking-for-broken-hyperlinks-in-excel
' Author is https://stackoverflow.com/users/2008576/tbur

Sub Audit_WorkSheet_For_Broken_Links()

If MsgBox("Is the Active Sheet a Sheet with Hyperlinks You Would Like to Check?", vbOKCancel) = vbCancel Then

    Exit Sub

End If

On Error Resume Next
For Each alink In Cells.Hyperlinks
    strURL = alink.Address

    If Left(strURL, 4) <> "http" Then
        strURL = ThisWorkbook.BuiltinDocumentProperties("Hyperlink Base") & strURL
    End If

    Application.StatusBar = "Testing Link: " & strURL
    Set objhttp = CreateObject("MSXML2.XMLHTTP")
    objhttp.Open "HEAD", strURL, False
    objhttp.Send

    If objhttp.statustext <> "OK" Then

        alink.Parent.Interior.Color = 255
    End If

Next alink
Application.StatusBar = False
On Error GoTo 0
MsgBox ("Checking Complete!" & vbCrLf & vbCrLf & "Cells With Broken or Suspect Links are Highlighted in RED.")

End Sub
