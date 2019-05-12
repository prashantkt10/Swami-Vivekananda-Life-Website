<% @ Language=VBScript %>
<% Option Explicit
response.buffer=true%>
<% dim num
num = Request.ServerVariables("SCRIPT_NAME")
application.lock
if isEmpty(application(num)) Then
application(num)=0
END IF
application(num)=application(num)+1
application.unlock
%>
<%
dim newmail, strfeed, ipadd, brow, now, msgbody, date, strnam, stremail, strcom, strsug
brow = Request.servervariables("HTTP_USER_AGENT")
ipadd= Request.ServerVariables("REMOTE_ADDR")
date = time
strfeed = request.cookies("Feed")("Yes")
if (Len(strfeed) > 0) then
response.write "<font color=red><b>You have already given us feedback! You can give it again after a week!</b></font>"
response.redirect "index.asp"
end if
strnam = Trim(Request.Form("name"))
stremail = Trim(request.form("mail"))
strcom = Trim(request.form("com"))
strsug = Trim(request.form("sug"))
num = application(num)
msgbody = ("Suggestions: " & strsug & "Comments: " & strcom & vbcrlf & "PC info:" & "IP:" & ipadd & "Browser: " & brow & "Time:" & date & "Hits:" & num &"times!")
if (strnam <> "" and stremail <> "" and strcom <> "" and strsug <> "")then
Set Newmail = Server.CreateObject ("cdonts.newmail")
Newmail.From = stremail
Newmail.To = "iadsdir@yahoo.co.uk"
Newmail.Subject = "Feedback From:" & strnam
Newmail.Body = msgbody
Newmail.Send
Set Newmail = Nothing 
response.cookies("Feed")("Yes") = "yes"
response.cookies("Feed")("Yes").expires = DateAdd("d",7,Now())
response.redirect "index.asp"
End if 
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Sample Feedback Form : Gopi Aravind</title>
<script language="JavaScript1.2">
var bookmarkurl="http://www.codeproject.com"
var bookmarktitle="The Code Project - Free Source Code and Tutorials"
function addbookmark()
{
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
</script>
</head>

<body bgcolor="#FFFFFF">

<script language=JavaScript>
<!--


var message="";
///////////////////////////////////
function clickIE() {if (document.all) {(message);return false;}}
function clickNS(e) {if 
(document.layers||(document.getElementById&&!document.all)) {
if (e.which==2||e.which==3) {(message);return false;}}}
if (document.layers) 
{document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;}
else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;}

document.oncontextmenu=new Function("return false")
// --> 
</script>

<p align="left"><font size="7" color="#00CCFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; F</font>eedback
<font size="7" color="#00CCFF">F</font>orm</p>

<div align="right">

<table border="0" cellspacing="1" id="AutoNumber2">
<form action="<%= Request.ServerVariables("SCRIPT_Name") %>" method="POST">

  <tr>
 <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
 <td><font color="#808080"><code>Name:</code></font></td>
    <td width="46%" height="29">
<input type="text" name="name" id=userID size="20" style="border: 1px solid #00CCFF; "></td>
  </tr>
  <tr>
    <td width="1%" height="29">&nbsp;</td>
    <td width="1%" height="29"><font color="#808080"><code>Email Id:</code></font></td>
    <td width="46%" height="29">
<input type="text" name="mail" id=userID0 size="20" maxlength="15" style="border: 1px solid #00CCFF; "></td>
  </tr>
  <tr>
    <td width="1%" height="36">&nbsp;</td>
    <td width="1%" height="36"><font color="#808080"><code>YourComments:</code></font></td>
    <td width="46%" height="36">
    <textarea rows="2" name="com" cols="20" style="border: 1px solid #00CCFF"></textarea></td>
  </tr>
  <tr>
    <td width="1%" height="31">&nbsp;</td>
    <td width="1%" height="31"><font color="#808080"><code>Suggestions:</code></font></td>
    <td width="46%" height="31">
    <textarea rows="2" name="sug" cols="20" style="border: 1px solid #00CCFF; "></textarea></td>
    </tr>
  <tr>
    <td width="1%" height="31">&nbsp;</td>
    <td width="1%" height="31">&nbsp;&nbsp;&nbsp;&nbsp; </td>
    <td width="46%" height="31">
<input type=submit  value="Send!" name=Submit style="border-left: 2px solid #003399; border-right-width: 1; border-top-width: 1; border-bottom: 2px solid #003399"> 
<input type="reset" value="Clear" id=reset name=reset style="border-left: 2px solid #003399; border-bottom: 2px solid #003399"></td>
  </tr>
</form>
</table>

</div>

<p>By: <font color="#808080">Gopi Aravind</font></p>
<p>Go here:<code><font color="#FF0000">
<a style="text-decoration: none" href="http://www.trollzsoft.com/feedback.asp">
http://www.trollzsoft.com/feedback.asp</a></font></code><a href="http://www.trollzsoft.com" style="text-decoration: none"> </a>
for an ever lasting sample!</p>

<p>Mail me at:
<a href="mailto:webmaster@trollzsoft.com">webmaster@trollzsoft.com</a> </p>
<p align="right">
<a href="javascript:addbookmark()"><font color="#808080">Bookmark 
www.codeproject.com</font></a><p align="right">
<font face="Trebuchet MS" color="#339933">
<a class="chlnk" style="cursor:hand" HREF onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.codeproject.com');"><u>
<font color="#808080">Set  as Codeproject.com as HomePage</font></u></a></font><hr color="#000000" width="89%">
<p align="center"><font color="#C0C0C0"><code>Copyright © 2003-2004 Gopi Aravind. 
This code is for free distribution only.</code></font></p>
<p align="center"><font color="#C0C0C0"><code>&nbsp;All Rights Reserved.</code></font></p>


<p align="center"><code><font color="#C0C0C0">Contact the webmaster:
<a href="mailto:webmaster@trollzsoft.com?subject=From the Homepage!">Gopi 
Aravind</a></font></code></p>

</body></html>