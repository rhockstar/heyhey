<%
'====================================================
'                                                   =
'               Forms To Go 4.5.4                   =
'            http://www.bebosoft.com/               =
'                                                   =
'====================================================
'                 UNREGISTERED COPY                 =
'              Forms To Go is shareware             =
'            Please register the software           =
'              http://www.bebosoft.com/             =
'====================================================




CONST kOptional = true
CONST kMandatory = false

CONST kStringRangeFrom = 1
CONST kStringRangeTo = 2
CONST kStringRangeBetween = 3
CONST kYes = "yes"
CONST kNo = "no"

CONST kNumberRangeFrom = 1
CONST kNumberRangeTo = 2
CONST kNumberRangeBetween = 3




'====================================================
' Function: FilterControlChars                      =
'====================================================

Function FilterCchar(TextToFilter)

 Dim regEx

 Set regEx = New RegExp
  
 regEx.Global = true
 regEx.IgnoreCase = true
 regEx.Pattern ="[\x00-\x1F]"

 filterCchar = regEx.Replace(TextToFilter, "")

End Function

'====================================================
' Function: SQLQuoteReplace                         =
'====================================================

Function SQLQuoteReplace(FieldValue)

 SQLQuoteReplace = Replace(FieldValue, "'", "''")

End Function

'====================================================
' Function: ValidateString                          =
'====================================================

Function check_string(field, low, high, mode, LimitAlpha, LimitNumbers, LimitEmptySpaces, LimitExtraChars, isOpt)

 check_string = false

 If LimitAlpha = kYes Then
  MyRegEx = "A-Za-z"
 End If
 
 If LimitNumbers = kYes Then
  MyRegEx = MyRegEx & "0-9"
 End If
 
 If LimitEmptySpaces = kYes Then
  MyRegEx = MyRegEx & " "
 End If

 If Len(LimitExtraChars) > 0 Then

  SpecialChars = "\,[,],-,$,.,*,(,),+,?,^,{,},|,/"
  SpecialCharsArray = Split(SpecialChars, ",")

  For cnt = 0 To UBound(SpecialCharsArray)
    LimitExtraChars = Replace(LimitExtraChars, SpecialCharsArray(cnt), "\" & SpecialCharsArray(cnt))
  Next
 
  MyRegEx = MyRegEx & LimitExtraChars
 End If

 Set regEx = New RegExp
 regEx.Pattern = "[^" & MyRegEx & "]"
 regEx.IgnoreCase = true

 If ( Len(field) > 0 ) And ( Len(MyRegEx) > 0 ) Then

  retVal = regEx.Test(field)
  
  If retVal Then
   Exit Function
  End If

 End If

 If ( (Len(field) = 0) and (isOpt = kOptional) ) Then

  check_string = true 

 Else

  If (mode = kStringRangeFrom) then
   If Len(field) >= low then
     check_string = true
   End If
  End If

  If (mode = kStringRangeTo) then
   If Len(field) <= high then
     check_string = true
   End If
  End If

  If (mode = kStringRangeBetween) then
   If Len(field) >= low and Len(field) <= high then
     check_string = true
   End If
  End If

 End If

End Function



'====================================================
' Function: ValidateEmail                           =
'====================================================

Function check_email(Email, isOpt)

 Dim regEx, retVal

 check_email = false

 If ( (Len(Email) = 0) and (isOpt = kOptional) ) Then

  check_email = true 

 Else
 
  Set regEx = New RegExp
  regEx.Pattern ="^([\w\!\#$\%\&\'\*\+\-\/\=\?\^\`{\|\}\~]+\.)*[\w\!\#$\%\&\'\*\+\-\/\=\?\^\`{\|\}\~]+@((((([a-z0-9]{1}[a-z0-9\-]{0,62}[a-z0-9]{1})|[a-z])\.)+[a-z]{2,6})|(\d{1,3}\.){3}\d{1,3}(\:\d{1,5})?)$"
  regEx.IgnoreCase = true

  retVal = regEx.Test(Email)

  If retVal Then
   check_email = true
  End If

 End If

End Function


'====================================================
' Function: ShowDate                                =
'====================================================

Function ShowDate(ftgdf)

 Dim FTGNow, FTGDay, FTGMonth, FTGYearS, FTGYearL, FTGHour, FTGMinute, FTGSecond, AMPM
 
 FTGNow = Now

 FTGDay = CStr(Day(FTGNow))
 FTGMonth = CStr(Month(FTGNow))
 FTGYearS = Right(CStr(Year(FTGNow)), 2)
 FTGYearL = CStr(Year(FTGNow))
 FTGHour = CStr(Hour(FTGNow))
 FTGMinute = CStr(Minute(FTGNow))
 FTGSecond = CStr(Second(FTGNow))

 If FTGDay < 10 Then FTGDay = "0" & FTGDay
 If FTGMonth < 10 Then FTGMonth = "0" & FTGMonth
 If FTGHour < 10 Then FTGHour = "0" & FTGHour
 If FTGMinute < 10 Then FTGMinute = "0" & FTGMinute
 If FTGSecond < 10 Then FTGSecond = "0" & FTGSecond

 If ftgdf = 1 Then ShowDate = FTGMonth & "/" & FTGDay & "/" & FTGYearS
 If ftgdf = 2 Then ShowDate = FTGDay & "/" & FTGMonth & "/" & FTGYearS
 If ftgdf = 3 Then ShowDate = FTGDay & "/" & FTGMonth & "/" & FTGYearL
 If ftgdf = 4 Then ShowDate = FTGYearL & "-" & FTGMonth & "-" & FTGDay
 If ftgdf = 6 Then ShowDate = FTGHour & ":" & FTGMinute & ":" & FTGSecond
 If ftgdf = 7 Then ShowDate = FTGYearL & "-" & FTGMonth & "-" & FTGDay & " " & FTGHour & ":" & FTGMinute & ":" & FTGSecond

 If ftgdf = 5 Then

  AMPM = "AM"
  FTGHour = Hour(FTGNow)
  If FTGHour > 12 Then
   FTGHour = FTGHour - 12
   AMPM = "PM"
   If FTGHour < 10 Then FTGHour = "0" & FTGHour
  ElseIf FTGHour = 12 Then
   AMPM = "PM"
  ElseIf FTGHour < 10 Then
   FTGHour = "0" & FTGHour
  End If
  
  ShowDate = FTGHour & ":" & FTGMinute & ":" & FTGSecond & " " & AMPM

 End If

End Function



Dim ClientIP

if Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "" then
 ClientIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
else
 ClientIP = Request.ServerVariables("REMOTE_ADDR")
end if

Dim objCDOSYSMail

Dim objCDOSYSCnfg
Set objCDOSYSCnfg = Server.CreateObject("CDO.Configuration")
FTGName = request.form("Name")
FTGCompany = request.form("Company")
FTGAddress = request.form("Address")
FTGAddress2 = request.form("Address2")
FTGCity = request.form("City")
FTGState = request.form("State")
FTGZip = request.form("Zip")
FTGPhone = request.form("Phone")
FTGEmail = request.form("Email")
FTGTypeBusiness = request.form("TypeBusiness")
FTGNumberEmployees = request.form("NumberEmployees")
FTGPayrollFreq = request.form("PayrollFreq")
FTGPayrollProcess = request.form("PayrollProcess")
FTGDirectDeposit = request.form("DirectDeposit")
FTGTaxFiling = request.form("TaxFiling")
FTG401 = request.form("401")
FTGNewHire = request.form("NewHire")
FTGRetirement = request.form("Retirement")
FTGWorkersComp = request.form("WorkersComp")
FTGTimeClock = request.form("TimeClock")
FTGInternetPayroll = request.form("InternetPayroll")
FTGAdditionalInfo = request.form("AdditionalInfo")
FTGSubmit = request.form("Submit")
FTGReset = request.form("Reset")

' Fields Validations

validationFailed = false

Dim FTGErrorMessage
Set FTGErrorMessage = Server.CreateObject("Scripting.Dictionary")

If (not check_email(FTGEmail, kMandatory)) Then
 validationFailed = true
 FTGErrorMessage.Add "Email", "Please enter a valid email address"
End If

If (not check_string(FTGAdditionalInfo, 0, 300, kStringRangeBetween, kNo, kNo, kNo, "", kMandatory)) Then
 validationFailed = true
 FTGErrorMessage.Add "AdditionalInfo", "300 charachter max"
End If


'====================================================
' Code: ErrorMessage                                =
'====================================================

If (validationFailed = true) Then

 ErrorPage = "<html><head><meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" /><title>Error</title></head><body>Errors found: <!--VALIDATIONERROR--></body></html>"


 dictItems = FTGErrorMessage.Items

 For cnt = 0 To FTGErrorMessage.Count - 1
  ErrorList = ErrorList & dictItems( cnt ) & "<br />"
 Next

 ErrorPage = Replace(ErrorPage, "<!--VALIDATIONERROR-->", ErrorList)


 Response.Write ErrorPage
 Response.End

End If



' Owner Email: cdosys
objCDOSYSCnfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "hybrid.gototransport.com"
objCDOSYSCnfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objCDOSYSCnfg.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCDOSYSCnfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objCDOSYSCnfg.Fields.Update

emailSubject = FilterCchar("Calgary Services, LLC information request")
emailBodyText = "Name : " & FTGName & "" & vbCrLf _
 & "Company : " & FTGCompany & "" & vbCrLf _
 & "Address : " & FTGAddress & "" & vbCrLf _
 & "Address2 : " & FTGAddress2 & "" & vbCrLf _
 & "City : " & FTGCity & "" & vbCrLf _
 & "State : " & FTGState & "" & vbCrLf _
 & "Zip : " & FTGZip & "" & vbCrLf _
 & "Phone : " & FTGPhone & "" & vbCrLf _
 & "Email : " & FTGEmail & "" & vbCrLf _
 & "Type Business : " & FTGTypeBusiness & "" & vbCrLf _
 & "Number Employees : " & FTGNumberEmployees & "" & vbCrLf _
 & "Payroll Freq : " & FTGPayrollFreq & "" & vbCrLf _
 & "Payroll Process : " & FTGPayrollProcess & "" & vbCrLf _
 & "Direct Deposit : " & FTGDirectDeposit & "" & vbCrLf _
 & "Tax Filing : " & FTGTaxFiling & "" & vbCrLf _
 & "401 : " & FTG401 & "" & vbCrLf _
 & "New Hire : " & FTGNewHire & "" & vbCrLf _
 & "Retirement : " & FTGRetirement & "" & vbCrLf _
 & "Workers Comp : " & FTGWorkersComp & "" & vbCrLf _
 & "Time Clock : " & FTGTimeClock & "" & vbCrLf _
 & "Internet Payroll : " & FTGInternetPayroll & "" & vbCrLf _
 & "Additional Info : " & FTGAdditionalInfo & "" & vbCrLf _
 & "Submit : " & FTGSubmit & "" & vbCrLf _
 & "Reset : " & FTGReset & "" & vbCrLf _
 & "" & vbCrLf _
 & ""

' Owner Email: cdosys
Set objCDOSYSMail = Server.CreateObject("CDO.Message")
objCDOSYSMail.Configuration = objCDOSYSCnfg

emailFrom = FilterCchar(FTGEmail)
objCDOSYSMail.To = "CalgaryServices Webform submit <HR@calgaryservices.biz>"
objCDOSYSMail.From = emailFrom
objCDOSYSMail.Subject = emailSubject
objCDOSYSMail.TextBody = emailBodyText
objCDOSYSMail.BodyPart.Charset = "UTF-8"

objCDOSYSMail.Send

'====================================================
' Code: SuccessMessage                              =
'====================================================

SuccessPage = "<html><head><meta http-equiv=""content-type"" content=""text/html; charset=utf-8"" /><title>Success</title></head><body>Form submitted successfully. It will be reviewed soon.</body></html>"

Response.Write SuccessPage



%>