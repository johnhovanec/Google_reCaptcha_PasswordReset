<%
	Option Explicit

	ToHTTPS
%>

<!-- #Include virtual = "/Main/Include/CommonFunc.asp" -->

<%
'	On Error Resume Next
	Dim strMedia
	Dim strCurrentMajorCategoryName	' current category's name
	Dim strCurrentSubCategoryName	' current sub-category's name
	Dim fColor						' leftinfoType1 each Section text color


	strCurrentLocation = "<font class=""st71""><a href = ""/"">Home</a></font> > <font class=""st7"">Forgot Password</font>"
%>

<html>
	<head>
		<title> <%=Application("PageTitle")%> </title>
		<script LANGUAGE="JavaScript" SRC="/Main/Js/CheckNorm.js"></script>
		<script language="JavaScript" src="/Main/Js/ValidateEmail.js"></script>
		<link href="/Main/CSS/DDStyle.css" type="text/css" rel="stylesheet">

		<script src='https://www.google.com/recaptcha/api.js'></script>						<!-- load Google reCAPTCHA  JH 6-26-17 -->
	</head>
	<body bgColor="#ffffff" style="margin:0" Onload="document.all.Email.focus()">
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
														<b><font class="MainTitle">Forgot Password?</font></b>
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
														<table cellSpacing="0" cellPadding="0" width="100%" border="0" ID="Table4">
															<tr>
																<td class="PageIntro">
																	Enter the email address for your account, select the checkbox to confirm you're not a robot, and click on the "Submit" button.  A temporary link will be sent to the submitted email address which will allow you to set a new password.
																</td>
															</tr>
															<tr>
																<td height="10">
																</td>
															</tr>
															<tr>
																<td>
																	<form name="FindPasswordForm" method="post" action="FindPwdOK.asp" OnSubmit="javascript:return CheckForm(this);">
																		<table width="600" cellSpacing="0" cellPadding="2" border="0" ID="Table5" align="center">
																			<tr>
																				<td height="20"></td>
																			</tr>
																			<tr>
																				<td>
																					<ul>
																						<img src="/images/dot2.gif" border="0" align="" WIDTH="4" HEIGHT="4"><font class="st6"> Email Address:</font>
																						&nbsp;&nbsp;<input type="text" name="Email" size="30" maxlength="32">
																					</ul>
																				</td>
																			</tr>
																			<tr>
																				<td>
																					<!-- add Google CAPTCHA widget --> 
																					<div class="g-recaptcha" data-sitekey="############-####################"></div>  
																				</td>
																			</tr>
																			<tr>
																				<td align="center"><input type="image" src="/images/btn_Submit.gif" border="0" WIDTH="60" HEIGHT="28">
																				</td>
																			</tr>
																		</table>
																		<script LANGUAGE="JavaScript">
																			function CheckForm( f )
																			{
																				if ( Checknorm( f.Email, 'Email Address') ) return false;
																				if ( ValidateEmail( f.Email, "Email Address" ) ) return false;
																			}
																		</script>
																	</form>
																</td>
															</tr>
														</table>
													</td>
													<td width="50">&nbsp;</td>
												</tr>
											</table>
											<tr>
												<td bgColor="#cccccc" height="1"></td>
											</tr>
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

<%
	' check error
	CheckError ""
%>
