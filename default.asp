<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta name="keywords" content="wine,Christopher">
<title>Christopher Stewart Wine &amp; Spirits</title>

<link REL="stylesheet" TYPE="text/css" HREF="CSS/hpCss.css">

<style TYPE="text/css">
	Body { margin-left:0 ; margin-right:0 ;margin-bottom:0; margin-top:0;
	}

    .copy-right{
           font-family:verdana;
            font-size:12px;
            color:white;
    }
</style>


<script language="JavaScript1.2" TYPE="TEXT/JAVASCRIPT">
	var topheight=58;
	var middleheight=558;
	var bottomheight=0;
	function submitForm()
	{
		alert("Login failed... Sorry, the username you entered does not exist.");
	}    
	
	

	function MouseOver(oImg,imgID)
	{
			if(imgID==0)		
				oImg.src = "images/register_button-down.jpg";
			else
				oImg.src = "images/enter-button-down.jpg";
	}
	function MouseLeave(oImg,imgID)
	{
			if(imgID==0)		
				oImg.src = "images/register_button-up.jpg";
			else
				oImg.src = "images/enter-button-up.jpg";

	}

	function setFocus()
	{
		document.getElementById("username").focus();
	}	
	
</script>
<!-- #include file="common.asp"-->

</head>
<%
	dim isIE
	isIE = 0
	If InStr(1, Request.ServerVariables("HTTP_USER_AGENT"), "MSIE") then
		isIE = 1
	End If	

	dim strUserId 
	dim strTopSrc
	strUserID =request("user_id")
	strTopSrc = "top.asp?user_id=" + strUserID
	dim strPageID
	strPageID = request("pageID")
	dim strProID
	strProID = 	Request.Cookies("province")
	'rememberuser = GetCookie("checkuser");
	dim bgcolor
	'bgcolor="#D4DAB8"
	bgcolor="#F2ECDB"
'	strProID="" 'only for test, must be command after use

	dim hghtmid
	hghtmid = 558
%>


<body id="mainpage" scroll="auto" onload="setFocus()">

<form id="index" action="index.asp" method="post">




<table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="1b161a" height="100%">
<!--up part-->
<tr><td width="100%" height="20%">
<table border="0" cellspacing="0" cellpadding="0" width="100%">
<tr><td width="100%" align="center">
<img src="images/houseparties_graphic1.jpg" WIDTH="639" HEIGHT="523">
</td></tr>

<tr><td width="100%" align="center">
<!-- user name and password-->
<!--table border="0" cellspacing="0" cellpadding="0"><tr><td align="right" class="td_fontB" style="color:white" style="padding-top:0px;padding-right:5px"> user name: </td><td><input id="username" class="pswd" ></td></tr><tr style="padding-bottom:40px;padding-top:8px"><td align="right" class="td_fontB" style="color:white; padding-right:5px"> password: </td><td><input id="passwd" class="pswd"></td></tr></table-->


		<table border="0" cellspacing="0" cellpadding="0" align="center" width="512" height="50">
		<tr>
			<td background="images/selectprovince2.jpg" valign="top" style="padding-top:0px;padding-left:0px;background-position:center ;
           background-repeat :no-repeat;">&nbsp;</td></tr></table>

		<table border="0" cellspacing="0" cellpadding="0" align="center" width="512" height="50">
		<tr><td width="300" style="padding-left:100;padding-top:10" height="50">
		<table border="0" cellspacing="0" cellpadding="0" height="50" style="font-family:verdana;font-size:11">
<tr><td align="right" class="td_fontB" style="color:white" style="padding-top:0px;padding-right:5px"> user name: </td><td><input id="username" class="pswd"></td></tr>
<tr style="padding-bottom:40px;padding-top:8px"><td align="right" class="td_fontB" style="color:white; padding-right:5px"> password: </td><td><input id="passwd" class="pswd"></td></tr>
		</table>
		
		</td>
		
		<td align="left" valign="top" valign="bottom" width="*" style="padding-top:0px;padding-left:0px">
				<img style="CURSOR: hand;" onclick="submitForm()" src="images/enter-button-up.jpg" onmouseover="MouseOver(this,1)" onmouseout="MouseLeave(this,1)" WIDTH="71" HEIGHT="73">
				</td>
</tr>		</table>

</td></tr>

</table>

</td></tr>

<!--blue part-->
<tr><td height="80%" valign="top">

<table border="0" cellspacing="0" cellpadding="0" height="100%" width="100%" bgcolor="0e426a">

<tr><td height="1%" valign="top" align="center" style="padding-top:20px;padding-bottom:50px" valign="top">
<a style="text-decoration: none" style="border:0" href="mailto:webmaster@houseparties.com?subject=Apply for access to houseparties.com"><img style="border:0" src="images/register_button-up.jpg" style="CURSOR: hand;" onmouseover="MouseOver(this,0)" onmouseout="MouseLeave(this,0)" WIDTH="256" HEIGHT="58"></a>

</td></tr>

<tr><td height="1%" valign="top" class="td_font" align="center" style="BORDER-bottom:1px #4c698f solid;padding-bottom:5px"><font color="white" size="1pt">A &nbsp;&nbsp;C H R I S T O P H E R &nbsp;&nbsp;S T E W A R T &nbsp;&nbsp;S I T E</td></tr>
<tr><td height="%%" valign="top" style="padding-top:5px" colspan="2"><div align="center" class="copy-right" id="copy_right">
Copyright &copy; 2018
<a style="text-decoration: none" href="http://www.christopherstewartwineandspirits.com/" target="_blank"><font face="verdana" color="white">Christopher Stewart Wine &amp; Spirits Inc.</font></a> All rights 
reserved </font></div>
			
			</td></tr>

</table>

</td></tr>

</table>

</form>
    
    <script type="text/javascript">
        document.getElementById(copy_right).innerHTML = "2019";
var gaJsHost = (("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">

try {
var pageTracker = _gat._getTracker("UA-12806568-2");
pageTracker._trackPageview();
} catch(err) {}</script>
</body>

</html>
<!--<table border="0" cellspacing="0" cellpadding="0"><tr><td></td></tr></table>"1b161a"bgcolor="1b161a"bgcolor="0e426a"-->

