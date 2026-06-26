# Automation-R-R
My Automation of R&amp;R workflow
<img width="1024" height="1536" alt="1000031724" src="https://github.com/user-attachments/assets/b0bdec3b-ed81-4069-8b8c-3e4cef397a33" />
<img width="1536" height="1024" alt="1000033137" src="https://github.com/user-attachments/assets/290a4904-559f-4de3-9fd1-89c0664c750a" />

🎓 Learning & Development
Training completion certificates
Batch learning reminders
Learning nomination emails
Course completion acknowledgements
Leadership program graduation certificates
🏆 Rewards & Recognition
Spot Award certificates
Employee of the Month recognition
Service anniversary certificates
Team achievement recognitions
Manager appreciation awards
👥 HR Operations
Offer letters
Internship completion certificates
Experience letters
Appreciation letters
Promotion letters
Welcome emails for new hires
📊 Employee Engagement
Event participation certificates
Wellness challenge recognition
Quiz and competition certificates
Festival event acknowledgements
CSR participation certificates
🤝 Client & Business Operations
Thank-you certificates for clients
Vendor appreciation letters
Partner recognition communications
Workshop participation certificates
📈 Advanced Version (Next Level)
Instead of only generating certificates, the workflow can:
Read participant data from Excel/Forms
Generate certificate
Convert to PDF
Send personalized email
Post recognition in Viva Engage
Update SharePoint tracker
Create Power BI dashboard
Notify managers automatically
Capture engagement metrics



VBA code:

Option Explicit

Private Const PLACEHOLDER_NAME As String = "Ashish Joshi"
Private Const PLACEHOLDER_AWARD_LINE As String = _
    "The Stellar Performance Award Q1-2026"

Sub Run_RR_Final()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim pptApp As Object
    Dim pptPres As Object
    Dim outApp As Object
    Dim mail As Object

    Dim templatePath As String
    Dim outputFolder As String
    Dim yearText As String
    Dim sendMode As VbMsgBoxResult

    Dim lastRow As Long
    Dim i As Long

    Dim nameCol As Long
    Dim emailCol As Long
    Dim awardCol As Long
    Dim quarterCol As Long
    Dim managerCol As Long

    Dim empName As String
    Dim empEmail As String
    Dim award As String
    Dim quarter As String
    Dim managerEmail As String
    Dim awardLine As String
    Dim pdfPath As String
    Dim safeName As String

    Dim stepName As String
    Dim rowInfo As String

    ' Find required columns
    stepName = "Finding header columns"

    nameCol = FindHeaderColumn(ws, "Employee Name")
    emailCol = FindHeaderColumn(ws, "Email")
    awardCol = FindHeaderColumn(ws, "Award Category")
    quarterCol = FindHeaderColumn(ws, "Quarter")
    managerCol = FindHeaderColumn(ws, "Manager Email ID")

    If nameCol = 0 Or emailCol = 0 Or awardCol = 0 _
       Or quarterCol = 0 Or managerCol = 0 Then

        MsgBox "Required headers not found in row 1." & vbCrLf & _
               vbCrLf & _
               "Needed headers:" & vbCrLf & _
               "Employee Name" & vbCrLf & _
               "Email" & vbCrLf & _
               "Award Category" & vbCrLf & _
               "Quarter" & vbCrLf & _
               "Manager Email ID", vbCritical

        Exit Sub
    End If

    ' Select PowerPoint template
    stepName = "Selecting PowerPoint template"

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Certificate Template"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "PowerPoint Files", "*.pptx"

        If .Show <> -1 Then Exit Sub

        templatePath = .SelectedItems(1)
    End With

    ' Select output folder
    stepName = "Selecting output folder"

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Output Folder"

        If .Show <> -1 Then Exit Sub

        outputFolder = .SelectedItems(1)
    End With

    ' Ask year
    stepName = "Getting year"

    yearText = Trim$(InputBox( _
        "Enter Certificate Year (example: 2026)", _
        "Certificate Year", _
        "2026"))

    If yearText = "" Then Exit Sub

    ' Ask email mode
    stepName = "Choosing email mode"

    sendMode = MsgBox( _
        "Choose email mode:" & vbCrLf & vbCrLf & _
        "YES = Send emails automatically now" & vbCrLf & _
        "NO = Preview each email before sending" & vbCrLf & _
        "CANCEL = Stop", _
        vbYesNoCancel + vbQuestion, _
        "Email Mode")

    If sendMode = vbCancel Then Exit Sub

    ' Open PowerPoint
    stepName = "Opening PowerPoint"

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    ' Open Outlook
    stepName = "Opening Outlook"

    Set outApp = CreateObject("Outlook.Application")

    ' Find last row
    stepName = "Finding last data row"

    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No employee data found below row 1.", vbExclamation
        GoTo SafeExit
    End If

    ' Loop through rows
    For i = 2 To lastRow

        rowInfo = "Row " & i

        stepName = rowInfo & " - Reading Excel values"

        empName = SafeText(ws.Cells(i, nameCol).Value)
        empEmail = SafeText(ws.Cells(i, emailCol).Value)
        award = SafeText(ws.Cells(i, awardCol).Value)
        quarter = SafeText(ws.Cells(i, quarterCol).Value)
        managerEmail = SafeText(ws.Cells(i, managerCol).Value)

        If empName = "" Or empEmail = "" _
           Or award = "" Or quarter = "" Then
            GoTo NextRow
        End If

        stepName = rowInfo & " - Building award line"

        awardLine = award & " " & quarter & "-" & yearText

        stepName = rowInfo & " - Building PDF path"

        safeName = CleanFileName(empName)
        pdfPath = outputFolder & "\" & safeName & ".pdf"

        ' Open template
        stepName = rowInfo & " - Opening PowerPoint template"

        Set pptPres = pptApp.Presentations.Open(templatePath)

        DoEvents

        stepName = rowInfo & " - Replacing employee name"

        ReplaceTextInPresentation pptPres, _
            PLACEHOLDER_NAME, empName

        DoEvents

        stepName = rowInfo & " - Replacing award line"

        ReplaceTextInPresentation pptPres, _
            PLACEHOLDER_AWARD_LINE, awardLine

        DoEvents

        stepName = rowInfo & " - Saving PDF"

        pptPres.SaveAs CStr(pdfPath), 32

        DoEvents

        stepName = rowInfo & " - Closing PowerPoint file"

        pptPres.Close
        Set pptPres = Nothing

        DoEvents

        ' Create Outlook email
        stepName = rowInfo & " - Creating Outlook email"

        Set mail = outApp.CreateItem(0)

        stepName = rowInfo & " - Populating email"

        With mail

            .To = empEmail
            .CC = managerEmail

            .Subject = "Congratulations - " & awardLine

            .Body = "Dear " & empName & "," & vbCrLf & vbCrLf & _
                    "Congratulations on receiving the " & awardLine & "." & vbCrLf & _
                    "Please find your certificate attached." & vbCrLf & vbCrLf & _
                    "Keep up the great work!" & vbCrLf & vbCrLf & _
                    "Regards," & vbCrLf & _
                    "EXL Analytics"

            If Len(Dir$(pdfPath)) > 0 Then
                .Attachments.Add pdfPath
            Else
                MsgBox "PDF not found for " & empName & ":" & _
                       vbCrLf & pdfPath, vbExclamation
            End If

            stepName = rowInfo & " - Sending/previewing email"

            If sendMode = vbYes Then
                .Send
            Else
                .Display
            End If

        End With

        Set mail = Nothing

NextRow:
    Next i

    If sendMode = vbYes Then
        MsgBox "Done! Certificates created and emails sent.", vbInformation
    Else
        MsgBox "Done! Certificates created and email previews opened.", vbInformation
    End If

SafeExit:

    On Error Resume Next

    If Not pptPres Is Nothing Then pptPres.Close
    If Not pptApp Is Nothing Then pptApp.Quit

    Set pptPres = Nothing
    Set pptApp = Nothing
    Set outApp = Nothing
    Set mail = Nothing

    Exit Sub

ErrHandler:

    MsgBox "Error: " & Err.Description & vbCrLf & _
           "Step: " & stepName & vbCrLf & _
           rowInfo, vbCritical

    Resume SafeExit

End Sub

Function SafeText(ByVal v As Variant) As String

    On Error Resume Next

    If IsError(v) Then
        SafeText = ""
    ElseIf IsNull(v) Then
        SafeText = ""
    Else
        SafeText = Trim$(CStr(v))
    End If

End Function

Sub ReplaceTextInPresentation(ByVal pres As Object, _
                              ByVal findText As String, _
                              ByVal replaceText As String)

    Dim sld As Object
    Dim shp As Object

    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            ReplaceTextInShape shp, findText, replaceText
        Next shp
    Next sld

End Sub

Sub ReplaceTextInShape(ByVal shp As Object, _
                       ByVal findText As String, _
                       ByVal replaceText As String)

    Dim i As Long

    On Error Resume Next

    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            shp.TextFrame.TextRange.Replace _
                FindWhat:=findText, _
                ReplaceWhat:=replaceText
        End If
    End If

    If shp.Type = 6 Then
        For i = 1 To shp.GroupItems.Count
            ReplaceTextInShape shp.GroupItems(i), _
                               findText, _
                               replaceText
        Next i
    End If

End Sub

Function FindHeaderColumn(ByVal ws As Worksheet, _
                          ByVal headerText As String) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim currentHeader As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol

        currentHeader = SafeText(ws.Cells(1, c).Value)

        If StrComp(currentHeader, headerText, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If

    Next c

    FindHeaderColumn = 0

End Function

Function CleanFileName(ByVal txt As String) As String

    txt = Replace(txt, "\", "_")
    txt = Replace(txt, "/", "_")
    txt = Replace(txt, ":", "_")
    txt = Replace(txt, "*", "_")
    txt = Replace(txt, "?", "_")
    txt = Replace(txt, """", "_")
    txt = Replace(txt, "<", "_")
    txt = Replace(txt, ">", "_")
    txt = Replace(txt, "|", "_")

    CleanFileName = txt

End Function
***********
