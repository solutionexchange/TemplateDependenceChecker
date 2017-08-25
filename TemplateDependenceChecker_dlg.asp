<% 
'Copyright (C) Stefan Buchali, UDG United Digital Group, www.udg.de
'
'This software is licensed under a
'Creative Commons GNU General Public License (http://creativecommons.org/licenses/GPL/2.0/)
'Some rights reserved.
'
'You should have received a copy of the GNU General Public License
'along with TemplateDependenceChecker.  If not, see http://www.gnu.org/licenses/.

Response.ContentType = "text/html"
Response.Charset = "utf-8"

RQLKey=Session("Sessionkey")
LoginGUID=Session("LoginGUID")

'Pre-defined Texts
pluginTitle = "Template-Verwendung in Tochterprojekten anzeigen" 'Show template usage in child projects
dlgFolder = "Ordner" 'Folder
dlgCK = "Content-Klasse" 'Content Class
dlgHeadline = "Achtung!" 'Attention!
dlgContent = "Bitte w&auml;hlen Sie eine Content-Klasse aus." 'Please choose a Content Class.
dlgAlsoMaster = "Auch die Instanzen des Masterprojekts zählen." 'Also take the master project's instances into account.
dlgPleaseWait = "Bitte warten" 'Please wait
dlgMessage = "Die &Uuml;berprüfung der abh&auml;ngigen Projekte kann einige Zeit in Anspruch nehmen!" 'Checking all child projects can take some minutes.
dlgOK = "OK"
dlgCancel = "Abbrechen" 'Cancel
dlgClose = "Schlie&szlig;en" 'Close
dlgContinue = false

'********************************
'nothing to edit below this point
'********************************

' XML-Verarbeitung per Microsoft-DOM vorbereiten
set XMLDoc = Server.CreateObject("MSXML2.DOMDocument")
XMLDoc.async = false
XMLDoc.validateOnParse = false
	
' RedDot Object fuer RQL-Zugriffe anlegen
set objIO = Server.CreateObject("OTWSMS.AspLayer.PageData")

' Variablendeklaration
Dim xmlSendDoc		' RQL-Anfrage, die zum Server geschickt wird
Dim ServerAnswer	' Antwort des Servers

'Daten aus Session lesen
if Session("TreeGuid")<>"" and Session("TreeParentGuid")<>"" Then

	TemplateGUID=Session("TreeGuid")
	TemplateFolderGUID=Session("TreeParentGuid")

	'Templateordner auslesen
	xmlSendDoc=	"<IODATA loginguid=""" & LoginGUID & """ sessionkey=""" & RQLKey & """>"&_
					"<TEMPLATEGROUPS action=""load"" />"&_
				"</IODATA>"
	ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
	XMLDoc.loadXML(ServerAnswer)

	set TemplatefolderList = XMLDoc.selectnodes("//GROUP")
	TemplateFolderGUID_is_a_Templatefolder = false
	for each Templatefolder in TemplatefolderList
		TemplateFolderGUID_is_a_Templatefolder = TemplateFolderGUID_is_a_Templatefolder or (Templatefolder.getAttribute("guid") = TemplateFolderGUID)
	next
	set TemplatefolderList = nothing

	if TemplateFolderGUID_is_a_Templatefolder Then
		TemplateFolderName = XMLDoc.selectSingleNode("//GROUP[@guid='" & TemplateFolderGUID & "']/@name").text
		dlgHeadline = dlgFolder & ": " & TemplateFolderName

		'Templates des Ordners auslesen
		xmlSendDoc=	"<IODATA loginguid=""" & LoginGUID & """ sessionkey=""" & RQLKey & """>"&_
						"<TEMPLATELIST action=""load"" folderguid=""" & TemplateFolderGUID & """/>"&_
					"</IODATA>"
		ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
		XMLDoc.loadXML(ServerAnswer)

		set TemplateList = XMLDoc.selectnodes("//TEMPLATE")
		TemplateGUID_is_a_Template = false
		for each Template in TemplateList
			TemplateGUID_is_a_Template = TemplateGUID_is_a_Template or (Template.getAttribute("guid") = TemplateGUID)
		next
		set TemplateList = nothing

		if TemplateGUID_is_a_Template Then
			TemplateName = XMLDoc.selectSingleNode("//TEMPLATE[@guid='" & TemplateGUID & "']/@name").text
			dlgContent = dlgCK & ": " & TemplateName
			dlgContinue = true
		end if
		
	end if

end if

' säubern
set XMLDoc = nothing
set objIO = nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" type="text/css" href="../../stylesheets/ioStyleSheet.css" />
<style type="text/css">
	div#messageDiv {
		color: #ff0000;
		font-weight: bold;
	}
	button {
		border: 1px solid black;
		background: #EEEEEE;
	}
</style>
<script type="text/javascript">
var submitOK=true;

function submitForm() {
	if(submitOK) {
		submitOK=false;
		document.getElementById("messageDiv").innerText='<%=dlgPleaseWait %>...';
		document.TemplateDependenceCheckerForm.submit();
		document.getElementById("cbMaster").disabled=true;
		document.getElementById("btn1").disabled=true;
		document.getElementById("btn2").disabled=true;
	}
}
</script>
<title><%=pluginTitle %></title>
</head>
<body link="navy" alink="navy" vlink="navy" bgcolor="#ffffff" background="../../icons/back5.gif">
<form name="TemplateDependenceCheckerForm" method="post" action="TemplateDependenceChecker_do.asp">
<% if dlgContinue then %>
<input type="hidden" name="TemplateFolderGUID" value="<%=TemplateFolderGUID %>" />
<input type="hidden" name="TemplateFolderName" value="<%=TemplateFolderName %>" />
<input type="hidden" name="TemplateGUID" value="<%=TemplateGUID %>" />
<input type="hidden" name="TemplateName" value="<%=TemplateName %>" />
<% end if %>
<table class="tdgrey" border="0" width="400" align="center" cellspacing="0" cellpadding="3">
<tr>
<td width="100%">
	<table class="tdgreylight" border="0" width="100%" cellspacing="0" cellpadding="1">
	<tr>
	<td width="100%" align="left" valign="top" height="50">
		<table border="0" width="100%">
		<tr><td class="titlebar" width="100%"><%=pluginTitle %></td></tr>
		</table>
	</td>
	</tr>
	<tr>
	<td width="100%" align="left" valign="top" height="80">
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
		<tr>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		<td align="left" valign="top" class="label" width="100%"><%= dlgHeadline %></td>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		</tr>
		<tr>
		<td height="5" colspan="3"></td>
		</tr>
		<tr>
		<td></td>
		<td><%=dlgContent %></td>
		<td></td>
		</tr>
<% if dlgContinue then %>
		<tr>
		<td height="15" colspan="3"></td>
		</tr>
		<tr>
		<td></td>
		<td><label><input type="checkbox" name="master" id="cbMaster" value="1" /> <%=dlgAlsoMaster %></label></td>
		<td></td>
		</tr>
		<tr>
		<td height="15" colspan="3"></td>
		</tr>
		<tr>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		<td align="left" valign="top" class="label" width="100%"><div id="messageDiv"><%=dlgMessage %></div></td>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		</tr>
		<tr>
		<td height="20" colspan="3"><img src="../../icons/transparent.gif" width="1" height="20" border="0" alt=""></td>
		</tr>
		<tr>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		<td align="right" valign="top"><button id="btn1" type="button" onclick="submitForm()"><%=dlgOK %></button>&nbsp;&nbsp;<button id="btn2" type="button" onclick="self.close()"><%=dlgCancel %></button></td>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		</tr>
<% else %>
		<tr>
		<td height="20" colspan="3"><img src="../../icons/transparent.gif" width="1" height="20" border="0" alt=""></td>
		</tr>
		<tr>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		<td align="right" valign="top"><button id="btn2" type="button" onclick="self.close()"><%=dlgClose %></button></td>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		</tr>
<% end if %>
		<tr>
		<td height="15" colspan="3"></td>
		</tr>
		</table>
	</td>
	</tr>
	</table>
</td>
</tr>
</table>
</form>

</body>
</html><script language="javascript" src="../../ioCheckEvent.js"></script>