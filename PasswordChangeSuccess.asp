<%
	Option Explicit

	ToHTTPS
%>

<!-- #Include virtual = "/Main/Include/CommonFunc.asp" -->

<%

   Session("Password") =  Session("NewPassword")
   Session("NewPassword") = "" 
   Session("OldPassword") = ""

	Dim strMedia					' Media Type
	Dim strCurrentMajorCategoryName	' current category's name
	Dim strCurrentSubCategoryName	' current sub-category's name
	Dim fColor						' leftinfoType1 each Section text color

	strCurrentLocation = "<font class=""st71""><a href = ""/"">Home</a></font> > <font class=""st7"">Register Success</font>"

	Dim strReturnUrl

	strReturnUrl = PathFilter(XSSFilter(Request.QueryString("ReturnUrl")))
	if strReturnUrl = "" then strReturnURL = "/default.asp"
				    
	Dim titleClass					' each Page title Color
	if strMedia = "Book" then
		titleClass = "BookTitle"
	elseif strMedia = "Music" then
		titleClass = "MusicTitle"
	else
		titleClass = "MainTitle"
	end if
%>

<!-- #Include virtual = "/Main/Include/ECommerce/RequestedVariable.asp" -->

<html>
	<head>
		<title> <%=Application("PageTitle")%> </title>
		<link href="/Main/CSS/DDStyle.css" type="text/css" rel="stylesheet">
	</head>
	<body bgColor="#ffffff" style="margin:0">
		<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
			<tr>
				<td height="119">
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
							<td vAlign="top" align="middle">
								<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" ID="Table1">
									<tr>
										<td colspan="5">
											<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" ID="Table2">
												<tr>
													<td height="40">
														<font class="<%=titleClass%>">Successful Password Change</font></b>
													</td>
												</tr>
												<tr>
													<td height="10"></td>
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
											<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" ID="Table3">
												<tr>
													<td>
														<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0" ID="Table4">
															<tr>
																<td>
																	<table cellSpacing="0" cellPadding="2" border="0" ID="Table5">
																		<tr>
																			<td class="st6">Congratulations. Your password has been changed.</td>
																		</tr>
																		<tr><td height="10"></td></tr>
																		<tr>
																			<td class="pageIntro">You may set preferences for browsing and products in "<a href="/Customer/">My Account</a>" 
																			(The icon is located on the right top of every page) to help with your shopping. 
																			</td>		
																		</tr>
																		<tr><td height="10"></td></tr>
																		<tr>
																			<td colspan="2" align="center" height="30">
																				<a href="<%=strReturnUrl%>"><img src="/images/btn_CONTINUE.gif" border="0" alt="continue" WIDTH="74" HEIGHT="28"></a>
																				<a href="https://<%=strServerName%>/Customer/"><img src="/images/btn_GoToMyAccount.gif" border="0" alt="GO TO MY PAGE" WIDTH="131" HEIGHT="28"> </a>
																			</td>
																		</tr>
																	</table>
																</td>
															</tr>
															<tr>
																<td height="30"></td>
															</tr>
														</table>
													</td>
												</tr>
											</table>
											<tr>
												<td height="20">
												</td>
											</tr>
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
				<td height="224">
					<!--Footer-->
<!-- #Include virtual = "/Main/Include/Footer.asp" -->
					<!--Footer-->
				</td>
			</tr>
		</table>
	</body>
</html>