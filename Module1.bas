Attribute VB_Name = "Module1"
Sub RunReport()

myDate = Format(Date, "mm dd yyyy")
Dim FileName As String
FileName = VBA.FileSystem.Dir("i:\Product Marketing\Dell\MyRewards\" & myDate & "\")
If FileName = VBA.Constants.vbNullString Then
    MkDir "i:\Product Marketing\Dell\MyRewards\" & myDate & "\"
Else
End If

A = Worksheets("Dell MyRewards Report").Cells(Rows.Count, 1).End(xlUp).Row
c = 2
x = 2
    Set Rng = Worksheets("Customers").Range("A:A")
    With Rng
        For Each Y In .Rows
        If Application.CountA(Y) > 0 Then
            x = x + 1
        End If
    Next
End With

Do While c < x - 1

    
    filepath = "i:\Product Marketing\Dell\MyRewards\" & myDate & "\" & Worksheets("Customers").Cells(c, 1).Value
    MyStr1 = Worksheets("Customers").Cells(c, 1).Value
    Email = Worksheets("Customers").Cells(c, 2).Value
        For i = 2 To A
            If Worksheets("Dell MyRewards Report").Cells(i, 29).Value = MyStr1 Then
                Worksheets("Dell MyRewards Report").Rows(i).Copy
                Worksheets("Report").Activate
                b = Worksheets("Report").Cells(Rows.Count, 1).End(xlUp).Row
                Worksheets("Report").Cells(b + 1, 1).Select
                ActiveSheet.Paste
            End If
    Worksheets("Report").Range("AE2") = Application.WorksheetFunction.Sum(Range("Y:Y"))
    Worksheets("Report").Range("H:L").EntireColumn.Hidden = True
    Worksheets("Report").Range("P:P").EntireColumn.Hidden = True
    Worksheets("Report").Range("R:V").EntireColumn.Hidden = True
    Worksheets("Report").Range("X:X").EntireColumn.Hidden = True
   
Next
 Worksheets("Report").Copy
 
  Application.DisplayAlerts = False
    
    With ActiveWorkbook
        .SaveAs FileName:=filepath, FileFormat:=51
            Dim wb As Workbook
            Dim FileExtStr As String
            Dim FileFormatNum As Long
            Dim TempFilePath As String
            Dim TempFileName As String
            Dim OutApp As Object
            Dim OutMail As Object
            Dim PromoSheet As String
            Dim SigString As String
            Dim Signature As String
            Dim strPath As String
            Dim objOutlookMsg As Object
            Dim strbody As String

            FileExtStr = ".xls"
            FileFormatNum = xlExcel8


            Set wb = ActiveWorkbook

            strbody = "<HTML><BODY>"
                strbody = strbody & "<font face =""Calibri"" size=""3"">" & "Good Day," & "<br>" & "<br>" _
                    & "Attached is the current MyRewards that you should be eligible to claim through Dell."

            strbody = strbody & "</BODY></HTML>"
            Set OutApp = CreateObject("Outlook.Application")
            Set objOutlookMsg = OutApp.CreateItem(olMailItem)
            Set objOutlookMsg = OutApp.CreateItem(0)
            With objOutlookMsg
                '.Display
        End With
        Signature = objOutlookMsg.Body
    
        With wb
        With objOutlookMsg
            .To = Email
            .Subject = "Your Dell and TD Synnex MyRewards Report"
            .HTMLBody = strbody & objOutlookMsg.HTMLBody
            .Attachments.Add (wb.FullName)
            
            .Send
        End With
    End With
            .Close SaveChanges = False
    End With
    Worksheets("Report").Activate
    Range("A2:A5000").Cells.EntireRow.Delete
    c = c + 1
    Loop

  
  
  
  
End Sub

