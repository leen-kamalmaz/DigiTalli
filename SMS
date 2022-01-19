Attribute VB_Name = "SendSMS"
Sub SendSMS()

'Authentication
APIKEY = ActiveSheet.Range("D4").Value

'Variables
toNumber = ActiveSheet.Range("C14").Value
BodyText = ActiveSheet.Range("D7").Value

'Use XML HHTP
Set Request = CreateObject("MSXML2.ServerXMLHTTP.6.0")

'Specify URL
Url = "https://gateway.sms77.io/api/sms?p=" & APIKEY & "&to=" & toNumber & "&text=" & BodyText & "&details=1"

'Open POST Request
Request.Open "POST", Url, False

'Request Header
Request.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

'Send Request
Request.send Url

Range("C14").Delete Shift:=xlUp


'Get response text (result)
MsgBox Request.responseText


End Sub
