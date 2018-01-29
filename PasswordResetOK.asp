<%
	Option Explicit
%>

<!-- #Include virtual = "/Main/Include/CommonFunc.asp" -->

<%
'	On Error Resume Next
   	ToHTTPS	

   	Dim objCustomer				' Customer object
	Dim intResult				' Result from sub call
	Dim strPassword				' Customer's new password
	Dim strEmailAddress 		' Customer's email address 
	Dim strReturnUrl

	strPassword = XSSFilter(Request.Form("Pwd"))
	strEmailAddress = XSSFilter(Request.Form("strTokenEmailAddress"))
	strReturnUrl = PathFilter(XSSFilter(Request.Form("ReturnUrl")))

	response.write"<br> strPassword = " & strPassword
	response.write"<br> strEmailAddress = " & strEmailAddress

	Set objCustomer = Server.CreateObject("RtlCustomer.clsCustomerBiz")

	' Call sub to set password
	intResult = Cint(objCustomer.SetPassword( strEmailAddress, strPassword ))

	response.write"<br> intResult = " & intResult

	CheckError ""
	
	If intResult = 0 then
		' if success
		Response.Redirect "/Customer/PasswordChangeSuccess.asp?ReturnUrl=" & strReturnUrl
	Else
		' Error with password change.
		Response.Redirect "/Customer/PasswordChangeError.asp?ReturnUrl=" & strReturnUrl
	End If

	Set objCustomer = Nothing
%>
