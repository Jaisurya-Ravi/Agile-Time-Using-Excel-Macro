VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents oSmtp As EASendMailObjLib.Mail
Attribute oSmtp.VB_VarHelpID = -1
Private CurrentEmailIsFinished As Boolean
Private HasErrorWithEmail As Boolean
Private ErrorDescription As String

Private Sub InitVariables()
    ' Initialize variables for email sending status and errors
    CurrentEmailIsFinished = True
    HasErrorWithEmail = False
    ErrorDescription = ""
End Sub

' Event handlers for email sending events
Private Sub oSmtp_OnAuthenticated()
    Application.StatusBar = "Authenticated"
End Sub

Private Sub oSmtp_OnClosed()
    CurrentEmailIsFinished = True
End Sub

Private Sub oSmtp_OnConnected()
    Application.StatusBar = "Connected"
End Sub

Private Sub oSmtp_OnError(ByVal lError As Long, ByVal ErrDescription As String)
    HasErrorWithEmail = True
    CurrentEmailIsFinished = True
    ErrorDescription = ErrDescription
End Sub

Private Sub oSmtp_OnSending(ByVal lSent As Long, ByVal lTotal As Long)
    Application.StatusBar = "Sending " & lSent & "/" & lTotal & " ..."
End Sub



Public Sub SendMailFromWorkBook(sheetName As String)
    ' Initialize Excel worksheet objects
    Dim oSheet1 As Worksheet
    Set oSheet1 = ThisWorkbook.Sheets(sheetName) ' Change the sheet name as needed
    
    'Navigate to the selected sheet
    Dim sheetNameNavigate As String
    sheetNameNavigate = sheetName ' Replace "Sheet2" with the name of the sheet you want to navigate to
    On Error Resume Next ' In case the sheet doesn't exist
    Sheets(sheetNameNavigate).Activate
    On Error GoTo 0 ' Reset error handling
    
    ' Get the last row and column with data in Time Sheet Details sheet
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = oSheet1.Cells(oSheet1.Rows.Count, 1).End(xlUp).row
    lastCol = oSheet1.Cells(1, oSheet1.Columns.Count).End(xlToLeft).Column
    
    
    
    
   ' Calculate the last row with data in column D
    Dim lastRow1 As Long
    lastRow1 = oSheet1.Cells(oSheet1.Rows.Count, "D").End(xlUp).row
    
    If lastRow1 < 2 Then
        MsgBox "No data found in column D!"
        Exit Sub
    End If
    ' Initialize a dictionary to store aggregated hours by activity type
    Dim activityDict As Object
    Set activityDict = CreateObject("Scripting.Dictionary")
    ' Populate dictionary with aggregated hours
    Dim i As Long
    For i = 2 To lastRow
        Dim activityType As String
        Dim hoursWorked As Double
        activityType = Trim(oSheet1.Cells(i, "C").Value)
        hoursWorked = oSheet1.Cells(i, "D").Value
        
        ' Check if activity type already exists in dictionary
        If activityDict.exists(activityType) Then
            ' Add hours worked to existing total
            activityDict(activityType) = activityDict(activityType) + hoursWorked
        Else
            ' Add activity type to dictionary with initial hours worked
            activityDict(activityType) = hoursWorked
        End If
    Next i
    ' Create a pie chart
    Dim chartObj As ChartObject
    Dim pieChart As Chart
    Set chartObj = oSheet1.ChartObjects.Add(Left:=100, Width:=375, Top:=50, Height:=225)
    Set pieChart = chartObj.Chart
    ' Populate arrays with data from dictionary
    Dim activityTypes() As String
    Dim hoursWorked1() As Double
    ReDim activityTypes(1 To activityDict.Count)
    ReDim hoursWorked1(1 To activityDict.Count)
    ' Fill arrays with data from dictionary
    Dim idx As Integer
    idx = 1
    For Each key In activityDict.Keys
        activityTypes(idx) = key
        hoursWorked1(idx) = activityDict(key)
        idx = idx + 1
    Next key
    ' Set source data for the pie chart based on arrays
    pieChart.SetSourceData Source:=oSheet1.Range("E1:F" & activityDict.Count), PlotBy:=xlColumns
    pieChart.SeriesCollection.NewSeries
    pieChart.SeriesCollection(1).XValues = activityTypes
    pieChart.SeriesCollection(1).Values = hoursWorked1
    pieChart.ChartType = xlPie
    
    
    
    
    
    ' Calculate total worked hours
    Dim totalWorkedHours As Double
    totalWorkedHours = Application.WorksheetFunction.Sum(oSheet1.Range("D2:D" & lastRow1))
    
    ' Add data labels to the pie chart
    pieChart.ApplyDataLabels
    
    ' Set the title of the pie chart dynamically
    pieChart.HasTitle = True
    pieChart.ChartTitle.Text = "Total Worked Hours: " & totalWorkedHours & " (In hours)"
    
    ' Remove the border around the pie chart
    pieChart.ChartArea.Border.LineStyle = xlNone
    
    ' Export the pie chart as an image
    Dim chartPath As String
    chartPath = ThisWorkbook.Path & "\PieChart.png"
    pieChart.Export Filename:=chartPath, FilterName:="PNG"
    

    
    
    
    
    ' Convert the image to Base64
    Dim base64Image As String
    base64Image = ConvertImageToBase64(chartPath)
    
    
    
    ' Create the HTML table
    Dim htmlTable As String
    htmlTable = "<style>" & _
            "table {" & _
            "  font-family: Calibri, Calibri;" & _
            "  border-collapse: collapse;" & _
            "  width: 100%;" & _
            "}" & _
            "td, th {" & _
            "  border: 1px solid #dddddd;" & _
            "  text-align: center;" & _
            "  padding: 2px;" & _
            "}" & _
            "tr:nth-child(even) {" & _
            "  background-color: #f2f2f2;" & _
            "}" & _
            "th {" & _
            "  background-color: #4CAF50;" & _
            "  color: white;" & _
            "  font-weight: bold;" & _
            "}" & _
            "</style>"
    htmlTable = htmlTable & "<table>"
    
    ' Populate the HTML table with data from Time Sheet Details sheet
    Dim row As Long, col As Long
    For row = 1 To lastRow
        htmlTable = htmlTable & "<tr>"
        For col = 1 To lastCol
            htmlTable = htmlTable & "<td>" & oSheet1.Cells(row, col).Value & "</td>"
        Next col
        htmlTable = htmlTable & "</tr>"
    Next row
    htmlTable = htmlTable & "</table>"
    
    ' Initialize Excel worksheet object for User Details
    Dim oSheet
    Set oSheet = ThisWorkbook.Sheets("User Details")

    ' Initialize Excel worksheet object for configuration settings
    ' Dim ws As Worksheet
    ' Set ws = ThisWorkbook.Worksheets("Configuration")

    ' Initialize variables for email details
    Dim sender, name, address, subject, bodyTemplate, body, bodyFormat, ServerAddr, ServerPort, username, password, ccaddress, Addnotes, empname
    bodyFormat = 0 'Text body format
    
    ' Please change sender address to yours
    sender = "timesheet.info@agile-labs.com"
    weekdate = oSheet1.Range("A2").Value
    formatweekdate = Format(weekdate, "dd-MMM-yyyy")
    subject = "Weekly Timesheet Detail for Employee " & oSheet.Range("B3").Value & " | " & formatweekdate
    Dim bodyText As String
    bodyText = "This is to provide you with the timesheet details of " & oSheet.Range("B3").Value & " for the week " & formatweekdate & ". Below are the details. "
    Addnotes = oSheet.Range("B8").Value
    empname = oSheet.Range("B3").Value
    
    
    ' Construct email body using a body template
    bodyTemplate = "Dear {name},<br><br>" & _
    bodyText & "<br><br>" & _
    Addnotes & "<br><br>" & _
    "<img src='data:image/png;base64," & base64Image & "'><br><br>" & _
    "{table} <br><br>" & _
    "Best Regards,<br>" & _
    empname & "<br><br>" & _
    "This email is automatically generated. Please do not reply to this message. *** Triggered on : " & Format(Now, "dd-mmm-yyyy hh:mm:ss")


    ' Sending Email
    Dim emailSent
    emailSent = 0
        
    ' Read email configuration settings from Configuration sheet
    name = oSheet.Range("B7").Value
    address = oSheet.Range("B6").Value
    ServerAddr = "smtp.office365.com"
    ServerPort = "587"
    username = "timesheet.info@agile-labs.com"
    password = "Agile!@#123"
    ccaddress = oSheet.Range("B5").Value
    
    
    
    ' Construct email body using the body template and HTML table
    body = Replace(bodyTemplate, "{name}", name)
    body = Replace(body, "{table}", htmlTable)
    
        ' Validate email configuration settings
        If Len(name) = 0 Then
        MsgBox "Manager Name is empty. Please enter your Manager Name."
        Exit Sub
        End If
    
        If Len(address) = 0 Then
        MsgBox "Your Email ID is empty. Please enter a valid Email ID."
        Exit Sub
        End If
        
        If Len(ServerAddr) = 0 Then
        MsgBox "Host Name or Server Address is empty. Please enter a valid Host Name or Server Address."
        Exit Sub
        End If
        
        If Len(ServerPort) = 0 Then
        MsgBox "Port is empty. Please enter a valid Port Number."
        Exit Sub
        End If
    
        If Len(username) = 0 Then
        MsgBox "Username is empty. Please enter a valid Outlook Email ID."
        Exit Sub
        End If
        
        If Len(password) = 0 Then
        MsgBox "Password is empty. Please enter a valid Outlook Email ID Password."
        Exit Sub
        End If
        
        If Len(ccaddress) = 0 Then
        MsgBox "Manager Email ID is empty. Please enter a valid Email ID."
        Exit Sub
        End If
        
        If Len(empname) = 0 Then
        MsgBox "Your Agile Time UserName is empty. Please enter a valid Email ID."
        Exit Sub
        End If
       
       ' Send email using SendMailTo function
        If Not SendMailTo(sender, name, address, subject, body, bodyFormat, ServerAddr, ServerPort, username, password, ccaddress, Addnotes, empname) Then
            Exit Sub
        End If


        emailSent = emailSent + 1
    'Next i

    Application.StatusBar = "Total " & emailSent & " email sent to " & address
    MsgBox "Total " & emailSent & " email sent to " & address

End Sub


Function ConvertImageToBase64(imagePath As String) As String
    Dim inputStream As Object
    Set inputStream = CreateObject("ADODB.Stream")
    inputStream.Type = 1 ' Binary
    inputStream.Open
    inputStream.LoadFromFile imagePath
    
    Dim base64Stream As Object
    Set base64Stream = CreateObject("MSXML2.DOMDocument").createElement("b64")
    base64Stream.DataType = "bin.base64"
    base64Stream.nodeTypedValue = inputStream.Read
    
    ConvertImageToBase64 = base64Stream.Text
End Function

Function SendMailTo(sender, name, address, subject, body, bodyFormat, ServerAddr, ServerPort, username, password, ccaddress, Addnotes, empname)
    ' Initialize EASendMail object
    Set oSmtp = New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt"

    ' Configure SMTP settings
    oSmtp.ServerAddr = ServerAddr
    oSmtp.username = username
    oSmtp.password = password

    ' Set server port, if 25 port doesn't work, try to use 587 port
    oSmtp.ServerPort = ServerPort

    ' Using TryTLS,
    ' If smtp server supports TLS, then TLS connection is used; otherwise, normal TCP connection is used.
    ' https://www.emailarchitect.net/easendmail/sdk/?ct=connecttype
    oSmtp.ConnectType = 4

    ' If your server is Exchange 2007 or later version, you can use EWS protocol.
    ' https://www.emailarchitect.net/easendmail/sdk/?ct=protocol
    ' Set Exchange Web Service Protocol - EWS - Exchange 2007/2010/2013/2016
    ' oSmtp.Protocol = 1

    ' Set email details
    oSmtp.FromAddr = sender
    oSmtp.AddRecipient name, address, 0
    oSmtp.AddRecipientEx "CC recipient<" & ccaddress & ">", 1
    oSmtp.subject = subject
    oSmtp.bodyFormat = 1
    oSmtp.bodyText = body
    
    ' Add primary recipient
    If address <> ccaddress Then
        oSmtp.AddRecipient name, address, 0 ' Add recipient with name and email address
    Else
        MsgBox "Email address and Manager Email address are the same. Please provide different addresses."
        SendMailTo = False ' Return False indicating email sending failed
        Exit Function
    End If

    ' You can add attachment like this:
    ' Add attachment from local disk
    ' If oSmtp.AddAttachment("d:\test.jpg") <> 0 Then
        ' Application.StatusBar = "Failed to add attachment with error:" & oSmtp.GetLastErrDescription()
        ' SendMailTo = False
        ' Exit Function
    'End If

    Application.DisplayStatusBar = True
    Application.StatusBar = "Connecting " & oSmtp.ServerAddr & " ..."

    ' Set asynchronous sending
    oSmtp.Asynchronous = 1
    InitVariables
    CurrentEmailIsFinished = False
    oSmtp.Sendmail
    
    ' Wait for the email sending, you can do other thing here
    Do While Not CurrentEmailIsFinished
        DoEvents
    Loop
    
    ' Check if there was an error while sending the email
    If HasErrorWithEmail Then
        Application.StatusBar = "Failed to send email to " & address & "; " & ErrorDescription
        SendMailTo = False
    Else
        Application.StatusBar = "Message to " & address & " has been submitted to server."
        SendMailTo = True
    End If

    Set oSmtp = Nothing
End Function


















  Sub GetDropdownValuesAndWriteToExcel()
        Dim driver As Selenium.WebDriver
        Set driver = New Selenium.WebDriver
        
        ' Start Chrome and navigate to the web page
        driver.Start "chrome"
        driver.Window.Maximize
        driver.Get "https://live.agilecloud.biz/agiletime/aspx/signin.aspx?agile_time_v2"
        
        
'        ' Send Alt + Spacebar, then N (combination to minimize window in Windows)
'        SendKeys "% ", True
'        driver.wait 200
'        SendKeys "N", True
'
'        ' Wait for a few seconds to let the window minimize
'        Application.wait Now + TimeValue("00:00:02")
    
        
        ' Set the worksheet where the data is located
        Dim ws1 As Worksheet
        Dim cellValue As Variant
        Set ws1 = ThisWorkbook.Sheets("User Details")

        ' Handling Username
        ' Read data from cell A1 (change the cell reference as needed)
        cellValueuser = ws1.Cells(3, 2).Value ' Row 1, Column 1 (A1)
        ' Enter username and click Next
        driver.FindElementByXPath("//input[@id='axUserName']").SendKeys cellValueuser
        driver.FindElementByXPath("//input[@name='btnNext']").Click
        ' Check if the username is empty and capture the error message
        Dim errorMessageuser As String
        On Error Resume Next
        errorMessageuser = driver.FindElementByXPath("//div[@class='toast-message']").Text
        On Error GoTo 0
        ' Print the error message to Excel if it exists
        If errorMessageuser <> "" Then
            ws1.Cells(3, 2 + 1).Value = errorMessageuser
            driver.Quit
            Application.StatusBar = "Error : " & errorMessageuser & vbCrLf & vbCrLf & "Try: Enter you Username in User Details sheet on the cell B3"
            MsgBox "Error : " & errorMessageuser & vbCrLf & vbCrLf & "Try: Enter you Username in User Details sheet on the cell B3"
            Else
            ws1.Cells(3, 2 + 1).ClearContents
        
        
        ' Handling Password
        ' Read data from the specified cell
        cellValue = ws1.Cells(4, 2).Value
        ' Enter password and click Submit
        driver.FindElementByXPath("//input[@id='axPassword']").SendKeys cellValue
        driver.FindElementByXPath("//input[@id='btnSubmitNew']").Click
        ' Check if the username is empty and capture the error message
        Dim errorMessagepass As String
        On Error Resume Next
        errorMessagepass = driver.FindElementByXPath("//div[@class='toast-message']").Text
        On Error GoTo 0
        ' Print the error message to Excel if it exists
        If errorMessagepass <> "" Then
            ws1.Cells(4, 2 + 1).Value = errorMessagepass
            driver.Quit
            Application.StatusBar = "Error : " & errorMessagepass & vbCrLf & vbCrLf & "Try: Enter you Password in User Details sheet on the cell B4"
            MsgBox "Error : " & errorMessagepass & vbCrLf & vbCrLf & "Try: Enter you Password in User Details sheet on the cell B4"
        Else
            ws1.Cells(4, 2 + 1).ClearContents
        ' Check if the password is incorrect and capture the error message
        Dim errorMessage As String
        On Error Resume Next
        errorMessage = driver.FindElementByXPath("//div[@class='toast-message']").Text
        On Error GoTo 0
        ' Print the error message to Excel if it exists
        If errorMessage <> "" Then
            ws1.Cells(4, 2 + 1).Value = errorMessage
            driver.Quit
            MsgBox "Error : " & errorMessage & "Try: Enter your valid Username and Password for your Agile Time"
        Else
            ws1.Cells(4, 2 + 1).ClearContents
            explicitwait driver, "//div[@class='loader-box-wrapper d-flex bg-white p-20 shadow rounded']"
    
    
    
        ' Click on menu items and switch to frame
        WaitForElementAndClick driver, "//span[text()='Time Sheet']", 5000
        driver.wait 2000
        WaitForElementAndClick driver, "//span[text()='Time Sheet Entry']", 5000
        driver.SwitchToFrame "middle1"
        explicitwait driver, "//div[@class='loader-box-wrapper d-flex bg-white p-20 shadow rounded']"
        Application.StatusBar = "Time Sheet Form Loaded"
        
        ' Get the specified Excel worksheet by name
        Dim xlWS As Worksheet
        Set xlWS = ThisWorkbook.Sheets("Master Details")
        
        
        ' Clearing Existing Master Data from Sheet
        Dim columnsToClear As Variant
        columnsToClear = Array(2, 3, 4, 5, 6)  ' Assuming Columns B, C, and D (change as needed)
        Dim rowIndex As Integer
        Dim colIndex As Variant
        ' Loop through each column in the array
        For Each colIndex In columnsToClear
            ' Clear all data in the current column starting from the second row
            For rowIndex = 2 To xlWS.Cells(xlWS.Rows.Count, colIndex).End(xlUp).row
                xlWS.Cells(rowIndex, colIndex).ClearContents
            Next rowIndex
        Next colIndex
        
        
        ' Dropdown 1
        ' Find the dropdown element by its XPath (replace with your actual XPath)
        driver.FindElementByXPath("//span[@aria-controls='select2-dwtask000F1-container']").Click
        driver.wait 500
        driver.FindElementByXPath("//a[@data-refresh='dwtask000F1']").Click
        driver.wait 3000
        Dim options1 As Object
        Set options1 = driver.FindElementsByXPath("//ul/li[@role='option']")
        ' Write the dropdown values to Excel starting from the specified row and column
        Dim rowIndex1 As Integer
        rowIndex1 = 2
        Dim colIndex1 As Integer
        colIndex1 = 2
        Dim optionIndex1 As Integer
        For optionIndex1 = 1 To options1.Count
            ' Write each option's text to Excel at the specified row and column
            xlWS.Cells(rowIndex1, colIndex1).Value = options1.Item(optionIndex1).Text
            rowIndex1 = rowIndex1 + 1 ' Move to the next row
        Next optionIndex1

        
'        ' driver.FindElementByXPath("(//button[@class='select2-selection__clear'])[1]").Click
'        'driver.wait 7000
'        driver.FindElementByXPath("//span[@aria-controls='select2-dwtask000F1-container']").Click
'        driver.FindElementByXPath("//a[@data-refresh='dwtask000F1']").Click
'        driver.wait 3000
'        driver.FindElementByXPath("(//ul/li[@role='option'])[2]").Click
'        driver.wait 500
'        driver.FindElementByXPath("//span[@aria-controls='select2-dwtask000F1-container']").Click
'        driver.FindElementByXPath("//a[@data-refresh='dwtask000F1']").Click
'        driver.wait 3000
'        driver.FindElementByXPath("//ul/li[text()='Adhoc']").Click
'        Dim Adhoc As String
'        Adhoc = driver.FindElementByXPath("//span[@id='select2-dwtask000F1-container']").Text
'        xlWS.Cells(2, 2).Value = Adhoc
        
        
        ' Dropdown 2
        ' Find the dropdown element by its XPath (replace with your actual XPath)
        driver.FindElementByXPath("//span[@aria-labelledby='select2-activitytype000F1-container']").Click
        driver.wait 500
        driver.FindElementByXPath("//a[@data-refresh='activitytype000F1']").Click
        driver.wait 3000
        Dim options2 As Object
        Set options2 = driver.FindElementsByXPath("//ul/li[@role='option']")
        ' Write the dropdown values to Excel starting from the specified row and column
        Dim rowIndex2 As Integer
        rowIndex2 = 2
        Dim colIndex2 As Integer
        colIndex2 = 3
        Dim optionIndex2 As Integer
        For optionIndex2 = 1 To options2.Count
            ' Write each option's text to Excel at the specified row and column
            xlWS.Cells(rowIndex2, colIndex2).Value = options2.Item(optionIndex2).Text
            rowIndex2 = rowIndex2 + 1 ' Move to the next row
        Next optionIndex2
        
        
        ' Dropdown 3
        ' Find the dropdown element by its XPath (replace with your actual XPath)
        driver.FindElementByXPath("//span[@aria-labelledby='select2-dwproject000F1-container']").Click
        driver.wait 500
        driver.FindElementByXPath("//a[@data-refresh='dwproject000F1']").Click
        driver.wait 3000
        Dim options3 As Object
        Set options3 = driver.FindElementsByXPath("//ul/li[@role='option']")
        ' Write the dropdown values to Excel starting from the specified row and column
        Dim rowIndex3 As Integer
        rowIndex3 = 2
        Dim colIndex3 As Integer
        colIndex3 = 4
        Dim optionIndex3 As Integer
        For optionIndex3 = 1 To options3.Count
            ' Write each option's text to Excel at the specified row and column
            xlWS.Cells(rowIndex3, colIndex3).Value = options3.Item(optionIndex3).Text
            rowIndex3 = rowIndex3 + 1 ' Move to the next row
        Next optionIndex3
        driver.FindElementByXPath("//input[@id='dwhoursspent000F1']").Click
        driver.wait 500
        
        
        ' Dropdown 4
        ' Find the dropdown element by its XPath (replace with your actual XPath)
        driver.FindElementByXPath("//span[@aria-labelledby='select2-partner000F1-container']").Click
        driver.wait 500
        driver.FindElementByXPath("//a[@data-refresh='partner000F1']").Click
        driver.wait 5000
        Dim options4 As Object
        Set options4 = driver.FindElementsByXPath("//ul/li[@role='option']")
        ' Write the dropdown values to Excel starting from the specified row and column
        Dim rowIndex4 As Integer
        rowIndex4 = 2
        Dim colIndex4 As Integer
        colIndex4 = 5
        Dim optionIndex4 As Integer
        For optionIndex4 = 1 To options4.Count
            ' Write each option's text to Excel at the specified row and column
            xlWS.Cells(rowIndex4, colIndex4).Value = options4.Item(optionIndex4).Text
            rowIndex4 = rowIndex4 + 1 ' Move to the next row
        Next optionIndex4
        
        ' Dropdown 5
        ' Find the dropdown element by its XPath (replace with your actual XPath)
        driver.FindElementByXPath("//span[@aria-labelledby='select2-customer000F1-container']").Click
        driver.wait 500
        driver.FindElementByXPath("//a[@data-refresh='customer000F1']").Click
        driver.wait 5000
        Dim options5 As Object
        Set options5 = driver.FindElementsByXPath("//ul/li[@role='option']")
        ' Write the dropdown values to Excel starting from the specified row and column
        Dim rowIndex5 As Integer
        rowIndex5 = 2
        Dim colIndex5 As Integer
        colIndex5 = 6
        Dim optionIndex5 As Integer
        For optionIndex5 = 1 To options5.Count
            ' Write each option's text to Excel at the specified row and column
            xlWS.Cells(rowIndex5, colIndex5).Value = options5.Item(optionIndex5).Text
            rowIndex5 = rowIndex5 + 1 ' Move to the next row
        Next optionIndex5
        driver.Quit
        Application.StatusBar = "Master Record Details are Updated Successfully!!"
        MsgBox "Master Record Details are Updated Successfully!!"
            
        End If
        End If
        End If
        driver.Quit
    End Sub












Sub BasicSeleniumVBAIntegration(sheetName As String)

        'Navigate to the selected sheet
        Dim sheetNameNavigate As String
        sheetNameNavigate = sheetName ' Replace "Sheet2" with the name of the sheet you want to navigate to
        On Error Resume Next ' In case the sheet doesn't exist
        Sheets(sheetNameNavigate).Activate
        On Error GoTo 0 ' Reset error handling
    
        Dim driver As Selenium.WebDriver
        Set driver = New Selenium.WebDriver
        
        ' Start Chrome and navigate to the web page
        driver.Start "chrome"
        driver.Window.Maximize
        
        driver.Get "https://live.agilecloud.biz/agiletime/aspx/signin.aspx?agile_time_v2"


        ' Set the worksheet where the data is located
        Dim ws1 As Worksheet
        Dim cellValue As Variant
        Set ws1 = ThisWorkbook.Sheets("User Details")
        
        
        ' Handling Username
        ' Read data from cell A1 (change the cell reference as needed)
        cellValueuser = ws1.Cells(3, 2).Value ' Row 1, Column 1 (A1)
        ' Enter username and click Next
        driver.FindElementByXPath("//input[@id='axUserName']").SendKeys cellValueuser
        driver.FindElementByXPath("//input[@name='btnNext']").Click
        ' Check if the username is empty and capture the error message
        Dim errorMessageuser As String
        On Error Resume Next
        errorMessageuser = driver.FindElementByXPath("//div[@class='toast-message']").Text
        On Error GoTo 0
        ' Print the error message to Excel if it exists
        If errorMessageuser <> "" Then
            ws1.Cells(3, 2 + 1).Value = errorMessageuser
            driver.Quit
            Application.StatusBar = "Error : " & errorMessageuser & vbCrLf & vbCrLf & "Try: Enter you Username in User Details sheet on the cell B3"
            MsgBox "Error : " & errorMessageuser & vbCrLf & vbCrLf & "Try: Enter you Username in User Details sheet on the cell B3"
        Else
            ws1.Cells(3, 2 + 1).ClearContents
        
        
        
        ' Handling Password
        ' Read data from the specified cell
        cellValue = ws1.Cells(4, 2).Value
        ' Enter password and click Submit
        driver.FindElementByXPath("//input[@id='axPassword']").SendKeys cellValue
        driver.FindElementByXPath("//input[@id='btnSubmitNew']").Click
        ' Check if the username is empty and capture the error message
        Dim errorMessagepass As String
        On Error Resume Next
        errorMessagepass = driver.FindElementByXPath("//div[@class='toast-message']").Text
        On Error GoTo 0
        ' Print the error message to Excel if it exists
        If errorMessagepass <> "" Then
            ws1.Cells(4, 2 + 1).Value = errorMessagepass
            driver.Quit
            Application.StatusBar = "Error : " & errorMessagepass & vbCrLf & vbCrLf & "Try: Enter you Password in User Details sheet on the cell B4"
            MsgBox "Error : " & errorMessagepass & vbCrLf & vbCrLf & "Try: Enter you Password in User Details sheet on the cell B4"
        Else
            ws1.Cells(4, 2 + 1).ClearContents
        ' Check if the password is incorrect and capture the error message
        Dim errorMessage As String
        On Error Resume Next
        errorMessage = driver.FindElementByXPath("//div[@class='toast-message']").Text
        On Error GoTo 0
        ' Print the error message to Excel if it exists
        If errorMessage <> "" Then
            ws1.Cells(4, 2 + 1).Value = errorMessage
            driver.Quit
            Application.StatusBar = "Error : " & errorMessage & "Try: Enter your valid Username and Password for your Agile Time"
            MsgBox "Error : " & errorMessage & "Try: Enter your valid Username and Password for your Agile Time"
        Else
            ws1.Cells(4, 2 + 1).ClearContents
            explicitwait driver, "//div[@class='loader-box-wrapper d-flex bg-white p-20 shadow rounded']"
    
    
    
        ' Click on menu items and switch to frame
        WaitForElementAndClick driver, "//span[text()='Time Sheet']", 5000
        driver.wait 2000
        WaitForElementAndClick driver, "//span[text()='Time Sheet Entry']", 5000
        driver.SwitchToFrame "middle1"
        explicitwait driver, "//div[@class='loader-box-wrapper d-flex bg-white p-20 shadow rounded']"

    
    ' Declare and set the worksheet variable
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName) ' Change "Sheet1" to your actual sheet name

    ' Get the used range in the Excel sheet (cells with data)
    Dim usedRange As Range
    Set usedRange = ws.usedRange
    
    ' Get the row count and column count of the used range
    Dim rowCount As Long
    Dim colCount As Long
    rowCount = usedRange.Rows.Count
    colCount = usedRange.Columns.Count
    
    ' Loop through rows and columns in the used range
    Dim row As Long
    Dim col As Long
    For row = 2 To rowCount
        ' Read data from the current row
        Dim cellValue1 As Variant
        Dim cellValue2 As Variant
        Dim cellValue3 As Variant
        Dim cellValue4 As Variant
        Dim cellValue5 As Variant
        cellValue1 = ws.Cells(row, 1).Value ' Assuming the first cell value is in column 1
        cellValue2 = ws.Cells(row, 2).Value
        cellValue3 = ws.Cells(row, 3).Value
        cellValue4 = ws.Cells(row, 4).Value
        cellValue5 = ws.Cells(row, 5).Value

        ' Check if the cell values are not empty
        If Not IsEmpty(cellValue1) And Not IsEmpty(cellValue2) Then
            ' Sending Date
            Dim formatteddate As String
            formatteddate = Format(cellValue1, "dd/mm/yyyy")
            driver.FindElementByXPath("//input[@id='dwdate000F1']").SendKeys formatteddate
            
            On Error Resume Next
            ' Sending Select Task Name
            driver.FindElementByXPath("//span[@aria-controls='select2-dwtask000F1-container']").Click
            driver.wait 1000
            driver.FindElementByXPath("//a[@data-refresh='dwtask000F1']").Click
            driver.wait 3000
            driver.FindElementByXPath("//li[text()='" & cellValue2 & "']").Click
            
            ' Sending Activity Type
            driver.FindElementByXPath("//span[@aria-labelledby='select2-activitytype000F1-container']").Click
            driver.wait 1000
            driver.FindElementByXPath("//a[@data-refresh='activitytype000F1']").Click
            driver.wait 3000
            driver.FindElementByXPath("//li[text()='" & cellValue3 & "']").Click
            
            ' Sending Hours Worked
            driver.FindElementByXPath("//input[@id='dwhoursspent000F1']").SendKeys cellValue4
            
''            Sending Partner Name
'            driver.FindElementByXPath("//span[@aria-labelledby='select2-partner000F1-container']").Click
'            driver.FindElementByXPath("//a[@data-refresh='partner000F1']").Click
'            driver.wait 3000
'            driver.FindElementByXPath("//li[text()='" & cellValue5 & "']").Click
            
''            Sending Customer Name
'            driver.FindElementByXPath("//span[@aria-labelledby='select2-customer000F1-container']").Click
'            driver.FindElementByXPath("//a[@data-refresh='customer000F1']").Click
'            driver.wait 3000
'            driver.FindElementByXPath("//li[text()='" & cellValue7 & "']").Click
            
            ' Sending Comments
            driver.FindElementByXPath("//a[@id='cke_18']").Click
            driver.FindElementByXPath("//div[@id='cke_1_contents']/textarea").SendKeys cellValue5
            
            ' submit Timesheet entry
            driver.FindElementByXPath("//a[@id='ftbtn_iSave']").Click
            driver.SwitchToDefaultContent
            
            ' Capture message (optional)
            Dim msg As String
            msg = driver.FindElementByXPath("//div[@class='toast-message']").Text
            Debug.Print msg
            On Error GoTo 0
            WaitForElementAndClick driver, "//span[text()='Time Sheet Entry']", 5000
            
            ' Write the message to the Excel sheet (optional)
            ws.Cells(row, colCount + 1).Value = msg ' Write to the end of the row
            driver.wait 3000
            driver.SwitchToFrame "middle1"
            
        End If
    Next row
    End If
    End If
    End If
    ' Wait and close the browser
    driver.wait 3000
    driver.Close
    Application.StatusBar = "Time Sheet Updated Successfully!!"
    MsgBox "Time Sheet Updated Successfully!!"
    
End Sub

    Sub explicitwait(driver As Object, locator As String)
        Dim i As Integer
        On Error GoTo Retry
        driver.FindElementByXPath (locator)
        Exit Sub
Retry:
        driver.wait 500
        i = i + 1
        If i = 20 Then
            On Error GoTo 0
            Exit Sub
        Else
            Resume
        End If
    End Sub
    
    Function WaitForElementAndClick(driver As Object, locator As String, timeout As Integer) As Boolean
        Dim startTime As Double
        startTime = Timer
        
        Do While Timer - startTime < timeout
            On Error Resume Next
            Dim element As Object
            Set element = driver.FindElementByXPath(locator)
            On Error GoTo 0
            
            If Not element Is Nothing Then
                element.Click
                WaitForElementAndClick = True
                Exit Function
            End If
            
            Application.wait Now + TimeValue("0:00:1") ' Wait for 1 second
        Loop
        
        WaitForElementAndClick = False
    End Function





Sub GoToSheetMasterDetails()
    Dim sheetName As String
    sheetName = "Master Details" ' Replace "Sheet2" with the name of the sheet you want to navigate to
    
    On Error Resume Next ' In case the sheet doesn't exist
    Sheets(sheetName).Activate
    On Error GoTo 0 ' Reset error handling
End Sub

Sub GoToSheetUserDetails()
    Dim sheetName As String
    sheetName = "User Details" ' Replace "Sheet2" with the name of the sheet you want to navigate to
    
    On Error Resume Next ' In case the sheet doesn't exist
    Sheets(sheetName).Activate
    On Error GoTo 0 ' Reset error handling
End Sub

Sub GoToSheetTimeSheetDetails()
    Dim sheetName As String
    sheetName = "Time Sheet Details" ' Replace "Sheet2" with the name of the sheet you want to navigate to
    
    On Error Resume Next ' In case the sheet doesn't exist
    Sheets(sheetName).Activate
    On Error GoTo 0 ' Reset error handling
End Sub

'Main Button action Starts here

Sub RunSheetSelectorForSendMail()
    ShowSheetSelectorForMail
End Sub


Sub RunSheetSelectorForUploadTimeSheetDetails()
    ShowSheetSelectorForUploadTimeSheetDetails
End Sub
