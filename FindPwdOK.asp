<%
	Option Explicit
%>

<!-- #Include virtual = "/Main/Include/CommonFunc.asp" -->

<%
	Dim objCustomer		' RtlCustomer object
	Dim strEmail		' Customer's Email address
	Dim strMedia
	Dim fColor
	Dim strToken 		' Token for password reset link
	Dim strVisitorID	' unique SessionID variable set in global.asa
	Dim strTitle		' Title for email message

	strEmail = Request.Form("Email")
	strEmail = replace(strEmail,"<","")
	strEmail = replace(strEmail,">","")
	strEmail = replace(StrEmail,"(","")
	strEmail = replace(StrEmail,")","")
	strEmail = replace(StrEmail,"'","")

'	On Error Resume Next

	Set objCustomer = Server.CreateObject("RtlCustomer.clsCustomerBiz")

	strVisitorID = XSSFilter(Session("VisitorID"))
	strToken = objCustomer.GetResetPasswordToken(strEmail, strVisitorID)

	Set objCustomer = Nothing 

	' Added to implement reCAPCHTA  JH 6-27-17 
	Dim recaptcha_secret
	recaptcha_secret = "########################-###########"  	' keep this secure!

	Dim sendstring
	sendstring = _
	   "https://www.google.com/recaptcha/api/siteverify?" & _ 
	   "secret=" & recaptcha_secret & _
	   "&response=" & request.form("g-recaptcha-response")

	Dim objXML
	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
	objXML.Open "GET", sendstring , false

	objXML.Send()
	If inStr(objXML.responseText,"true") then					' search for 'true' within the object returned from reCAPTCHA test
	        'Response.Write "<br> reCAPCHTA True"          
	Else
		%>
		<script type="text/javascript">
			alert("Please click the \"I'm Not A Robot\" checkbox before submitting your request.");
			history.go(-1);
	    </script>
    <%             
	End if

%>

<html>
	<head>
		<title> <%=Application("PageTitle")%> </title>
		<link href="/Main/CSS/DDStyle.css" type="text/css" rel="stylesheet">
	</head>
	<body bgColor="#ffffff" style="margin:0">
		<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
			<tr>
				<td colSpan="3" height="119">
					<!--BooksHeader-->
						<!-- #include virtual = "/Main/Include/MainTopMenu.asp" -->
					<!--BooksHeader-->
				</td>
			</tr>
			<tr>
				<td height="1" bgcolor="#9A9A9A" colspan="3"></td>
			</tr>
			<tr>
				<td height="25" bgcolor="#DADADA" colspan="3">
					<!-- #include virtual = "/Main/Include/Search.asp" -->
				</td>
			</tr>
			<tr>
				<td>
					<table cellSpacing="0" cellPadding="0" width="100%" border="0">
						<tr>
							<td vAlign="top" width="160">
								<table id="Table1" cellSpacing="0" cellPadding="0" width="160" border="0">
									<tr>
										<td rowspan="6"><img src="/images/blank5px.gif" width="3"></td>
										<td height="13"></td>
										<td rowspan="6"><img src="/images/blank5px.gif" width="3"></td>
									</tr>
									<tr>
										<td vAlign="top">
											<!--YourCart-->
												<!-- #include virtual = "/Main/Include/MiniCart.asp" -->
											<!--YourCart-->
										</td>
									</tr>
									<tr>
										<td height="5" colspan="3"></td>
									</tr>
									<tr>
										<td vAlign="top">
											<!--LeftInfoType1-->
												<!-- #include virtual = "/Main/Include/LeftInfoType1.asp" -->
											<!--LeftInfoType1-->
										</td>
									</tr>
								</table>
							</td>
							<td vAlign="top" align="top">
								<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" ID="Table1">
									<tr>
										<td colspan="5">
											<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" ID="Table2" height="100%">
												<tr>
													<td height="10"></td>
												</tr>
												<tr>
													<td>
														<b><font class="MainTitle">Password Assistance</font></b>
													</td>
												</tr>
												<tr>
													<td height="20"></td>
												</tr>
												<tr>
													<td height="10" bgcolor="#EDEDED"></td>
												</tr>
												<tr>
													<td height="10"></td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td>
											<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" ID="Table3" height="100%">
												<tr>
													<td>
														<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" ID="Table4">
															<tr>
																<td class="PageIntro">

<%

	' Email message title
	strTitle = "DaedalusBooks.com Account Information."


	' Test token for valid email address and the number of password reset links did not exceed the limit of 10 per 24hrs.
	If  InStr(strToken, "-1") Then
		'If strToken = -1 there are 10+ attempts within 24 hours. 
		Response.write "There seems to be a problem with your requested password reset link. If you have placed multiple requests and are unable to reset your password, please contact <a href=""/Main/Help/ContactUs.asp"">Customer Service</a> for assistance."
	ElseIf InStr(strToken, "-2") Then
		'If strToken = -2 that implies the user entered an email address that is not in our records, but we don't want to share that and we don't send an email.
		Response.Write "If the email address <font class=""st6"">" & strEmail & "</font> is associated with an account in our customer database, you will receive an email message with a link for you to set a new password. It should arrive in your email box shortly."
	Else
		'If here, a matching email account was found so we send the email
		SendSMTPMail "custserv@daedalusbooks.com", strEmail, strTitle, strBody, ""	
		Response.Write "If the email address <font class=""st6"">" & strEmail & "</font> is associated with an account in our customer database, you will receive an email message with a link for you to set a new password. It should arrive in your email box shortly."
	End if

%>
																</td>
															</tr>
															<tr>
																<td height="20">
																</td>
															</tr>
															<tr>
																<td align="center" height="20">
																	<a href="https://<%=strServerName%>/"><img src="/images/btn_CONTINUESHOPPINGbutton.gif" border="0" WIDTH="138" HEIGHT="28"></a>
																</td>
															</tr>
														</table>
													</td>
													<td width="50">&nbsp;</td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td height="50"></td>
									</tr>
								</table>
							</td>
							<td vAlign="top" width="10"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td colSpan="3" height="224">
					<!--Footer-->
						<!-- #Include virtual = "/Main/Include/Footer.asp" -->
					<!--Footer-->
				</td>
			</tr>
		</table>
	</body>
</html>