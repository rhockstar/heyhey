<%
' ----------------------------------------------------
' -----
' -----  Forms To Go v2.6.12 by Bebosoft, Inc.
' -----
' -----  http://www.bebosoft.com/
' -----
' ----------------------------------------------------
' -----
' -----                 UNREGISTERED COPY
' -----
' -----              Forms To Go is shareware
' -----            Please register the software
' -----
' ----------------------------------------------------

'----------
' Filter Control Characters

Function filterCchar(TextToFilter)

 Dim regEx

 Set regEx = New RegExp
  
 regEx.Global = true
 regEx.IgnoreCase = true
 regEx.Pattern ="[\x00-\x1F]"

 filterCchar = regEx.Replace(TextToFilter, "")

End Function
Dim ClientIP

if Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "" then
 ClientIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
else
 ClientIP = Request.ServerVariables("REMOTE_ADDR")
end if

Dim objCDOSYSMail
Set objCDOSYSMail = Server.CreateObject("CDO.Message")
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

' Redirect user to the error page

If (validationFailed = true) Then

 Response.Redirect "error.html"
 Response.End

End If

' Owner Email: cdosys

objCDOSYSCnfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "infoportal.expressleasing.local"
objCDOSYSCnfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objCDOSYSCnfg.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCDOSYSCnfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objCDOSYSCnfg.Fields.Update

objCDOSYSMail.Configuration = objCDOSYSCnfg
emailFrom = FilterCchar("quote@calgaryservices.biz")
emailSubject = FilterCchar("Calgary Services, LLC information request")
emailBodyText = "Name: " & FTGName & "" & vbCrLf _
 & "Company: " & FTGCompany & "" & vbCrLf _
 & "Address: " & FTGAddress & "" & vbCrLf _
 & "Address2: " & FTGAddress2 & "" & vbCrLf _
 & "City: " & FTGCity & "" & vbCrLf _
 & "State: " & FTGState & "" & vbCrLf _
 & "Zip: " & FTGZip & "" & vbCrLf _
 & "Phone: " & FTGPhone & "" & vbCrLf _
 & "Email: " & FTGEmail & "" & vbCrLf _
 & "TypeBusiness: " & FTGTypeBusiness & "" & vbCrLf _
 & "NumberEmployees: " & FTGNumberEmployees & "" & vbCrLf _
 & "PayrollFreq: " & FTGPayrollFreq & "" & vbCrLf _
 & "PayrollProcess: " & FTGPayrollProcess & "" & vbCrLf _
 & "DirectDeposit: " & FTGDirectDeposit & "" & vbCrLf _
 & "TaxFiling: " & FTGTaxFiling & "" & vbCrLf _
 & "401: " & FTG401 & "" & vbCrLf _
 & "NewHire: " & FTGNewHire & "" & vbCrLf _
 & "Retirement: " & FTGRetirement & "" & vbCrLf _
 & "WorkersComp: " & FTGWorkersComp & "" & vbCrLf _
 & "TimeClock: " & FTGTimeClock & "" & vbCrLf _
 & "InternetPayroll: " & FTGInternetPayroll & "" & vbCrLf _
 & "AdditionalInfo: " & FTGAdditionalInfo & "" & vbCrLf _
 & "" & vbCrLf _
 & "" & vbCrLf _
 & ""

objCDOSYSMail.To = "Human Resources <hr@calgaryservices.biz>"
objCDOSYSMail.From = emailFrom
objCDOSYSMail.Subject = emailSubject
objCDOSYSMail.TextBody = emailBodyText
objCDOSYSMail.BodyPart.Charset = "ISO-8859-1"
objCDOSYSMail.Send

' Redirect user to success page

Response.Redirect "http://www.calgaryservices.biz/success.html"


' End of ASP script
%>