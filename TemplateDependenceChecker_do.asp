<% 
'Copyright (C) 2009  Stefan Buchali, SF eBusiness GmbH, Herrenberg, Germany, www.sfe.de
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
UserName = ""
UserPW = ""

pluginTitle = "Template-Verwendung in Tochterprojekten anzeigen" 'Show template usage in child projects
dlgProject = "Projekt" 'Project
dlgUsedIn = "Verwendet in" 'Assigned to
dlgProjectVariants = "Projektvarianten" 'project variants
dlgInstances = "Instanzen" 'Instances
dlgError = "Fehler" 'Error
dlgContentClassNotFound = "Content-Klasse nicht gefunden" 'Content class not found
dlgFolderNotFound = "Ordner nicht gefunden" 'Folder not found
dlgNoRights = "Keine Rechte" 'No rights
dlgInsufficientRights = "Keine ausreichenden Admin-Rechte, um die Prüfung durchzuführen" 'Insufficient user rights to perform the task
dlgHeadline = "Folgende Templates werden verwendet" 'The following templates are in use
dlgClose = "Schlie&szlig;en" 'Close

'********************************
'nothing to edit below this point
'********************************

resultStr = ""

TemplateFolderGUID = Request.Form("TemplateFolderGUID")
TemplateFolderName = Request.Form("TemplateFolderName")
TemplateGUID = Request.Form("TemplateGUID")
TemplateName = Request.Form("TemplateName")


' XML-Verarbeitung per Microsoft-DOM vorbereiten
set XMLProjDoc = Server.CreateObject("MSXML2.DOMDocument")
XMLProjDoc.async = false
XMLProjDoc.validateOnParse = false

set XMLFoldersDoc = Server.CreateObject("MSXML2.DOMDocument")
XMLFoldersDoc.async = false
XMLFoldersDoc.validateOnParse = false

set XMLDoc = Server.CreateObject("MSXML2.DOMDocument")
XMLDoc.async = false
XMLDoc.validateOnParse = false
	
set XMLTemplatesDoc = Server.CreateObject("MSXML2.DOMDocument")
XMLTemplatesDoc.async = false
XMLTemplatesDoc.validateOnParse = false
	
' RedDot Object fuer RQL-Zugriffe anlegen
set objIO = Server.CreateObject("OTWSMS.AspLayer.PageData")

' Variablendeklaration
Dim xmlSendDoc		' RQL-Anfrage, die zum Server geschickt wird
Dim ServerAnswer	' Antwort des Servers


'Freigegebene Projekte auslesen
xmlSendDoc=	"<IODATA loginguid=""" & LoginGUID & """>"&_
				"<PROJECT sessionkey=""" & RQLKey & """>"&_
					"<SHAREDFOLDER action=""load"" guid=""" & TemplateFolderGUID & """ />"&_
				"</PROJECT>"&_
			"</IODATA>"
ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
XMLProjDoc.loadXML(ServerAnswer)

Set ProjekteList = XMLProjDoc.selectNodes("//PROJECT[@sharedrights='1']")



'Login als RQLAdmin
xmlSendDoc=	"<IODATA>"&_
				"<ADMINISTRATION action=""login"" name=""" & UserName & """ password=""" & UserPW & """ />"&_
			"</IODATA>"
ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
if InStr(ServerAnswer,"guid")>0 then
	XMLDoc.loadXML(ServerAnswer)
	RqlAdmLoginGUID = XMLDoc.selectsinglenode("/IODATA/LOGIN/@guid").text

	for each Projekt in ProjekteList
	
		resultStr = resultStr & "<hr/><p><b>" & dlgProject & ": " & Projekt.getAttribute("name") & "</b></p>"
		
		'In Projekt einchecken
		xmlSendDoc=	"<IODATA loginguid=""" & RqlAdmLoginGUID & """>"&_
						"<ADMINISTRATION action=""validate"" guid=""" & RqlAdmLoginGUID & """ useragent=""script"">"&_
							"<PROJECT guid=""" & Projekt.getAttribute("guid") & """/>"&_
						"</ADMINISTRATION>"&_
					"</IODATA>"
		ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
		if InStr(ServerAnswer,"key")>0 then
			XMLDoc.loadXML(ServerAnswer)
			RqlAdmSessionKey = XMLDoc.selectsinglenode("//SERVER/@key").text
			
			'Templateordner herausfinden
			xmlSendDoc=	"<IODATA loginguid=""" & RqlAdmLoginGUID & """ sessionkey=""" & RqlAdmSessionKey & """>"&_
							"<PROJECT>"&_
								"<FOLDERS action=""list"" foldertype=""1""/>"&_
							"</PROJECT>"&_
						"</IODATA>"
			ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
			XMLFoldersDoc.loadXML(ServerAnswer)
			
			Set SharedFoldersList = XMLFoldersDoc.selectNodes("//FOLDER")
			TemplateFolderToCheckGUID = ""
			for each SharedFolder in SharedFoldersList
				xmlSendDoc=	"<IODATA loginguid=""" & RqlAdmLoginGUID & """ sessionkey=""" & RqlAdmSessionKey & """>"&_
								"<PROJECT>"&_
									"<FOLDER action=""load"" guid=""" & SharedFolder.getAttribute("guid") & """/>"&_
								"</PROJECT>"&_
							"</IODATA>"
				ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
				XMLDoc.loadXML(ServerAnswer)
				if XMLDoc.selectNodes("//FOLDER/@linkedfolderguid").length = 1 then
					if XMLDoc.selectSingleNode("//FOLDER/@linkedfolderguid").text = TemplateFolderGUID then
						TemplateFolderToCheckGUID = XMLDoc.selectSingleNode("//FOLDER/@guid").text
					end if
				end if
			next
			Set SharedFoldersList = nothing
			
			if TemplateFolderToCheckGUID<>"" then
				'Content-Klassen auslesen
				xmlSendDoc=	"<IODATA loginguid=""" & RqlAdmLoginGUID & """ sessionkey=""" & RqlAdmSessionKey & """>"&_
								"<TEMPLATELIST action=""load"" folderguid=""" & TemplateFolderToCheckGUID & """/>"&_
							"</IODATA>"
				ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
				XMLDoc.loadXML(ServerAnswer)
				
				Set SharedTemplatesList = XMLDoc.selectNodes("//TEMPLATE")
				TemplateToCheckGUID = ""
				for each SharedTemplate in SharedTemplatesList
					if SharedTemplate.getAttribute("name") = TemplateName then
						TemplateToCheckGUID = SharedTemplate.getAttribute("guid")
					end if
				next				
				Set SharedTemplatesList = nothing
				
				if TemplateToCheckGUID<>"" then
				
					'Templates auslesen
					xmlSendDoc=	"<IODATA loginguid=""" & RqlAdmLoginGUID & """ sessionkey=""" & RqlAdmSessionKey & """>"&_
									"<PROJECT>"&_
										"<TEMPLATE guid=""" & TemplateToCheckGUID & """>"&_
											"<TEMPLATEVARIANTS action=""list""/>"&_
										"</TEMPLATE>"&_
									"</PROJECT>"&_
								"</IODATA>"
					ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
					XMLTemplatesDoc.loadXML(ServerAnswer)
					Set TemplateVariantList = XMLTemplatesDoc.selectNodes("//TEMPLATEVARIANT")
					
					'Projektvarianten-Zuweisung auslesen
					xmlSendDoc=	"<IODATA loginguid=""" & RqlAdmLoginGUID & """ sessionkey=""" & RqlAdmSessionKey & """>"&_
									"<PROJECT>"&_
										"<TEMPLATE guid=""" & TemplateToCheckGUID & """>"&_
											"<TEMPLATEVARIANTS action=""projectvariantslist""/>"&_
										"</TEMPLATE>"&_
									"</PROJECT>"&_
								"</IODATA>"
					ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
					XMLDoc.loadXML(ServerAnswer)

					resultStr = resultStr & "<p>"
					for each TemplateVariant in TemplateVariantList
						resultStr = resultStr & TemplateVariant.getAttribute("name") & ": "
						if XMLDoc.selectNodes("//TEMPLATEVARIANT[@guid='" & TemplateVariant.getAttribute("guid") & "']").length <> 0 then
							resultStr = resultStr & dlgUsedIn & " " & XMLDoc.selectNodes("//TEMPLATEVARIANT[@guid='" & TemplateVariant.getAttribute("guid") & "']").length & " " & dlgProjectVariants
						else
							resultStr = resultStr & "-"
						end if
						resultStr = resultStr & "<br/>"
					next
					resultStr = resultStr & "</p>"

					Set TemplateVariantList = nothing
					
					'Instanzen zählen
					xmlSendDoc=	"<IODATA loginguid=""" & RqlAdmLoginGUID & """ sessionkey=""" & RqlAdmSessionKey & """>"&_
									"<PAGE action=""search"" templateguid=""" & TemplateToCheckGUID & """ flags=""0"" maxrecords=""999999""/>"&_
								"</IODATA>"
					ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)
					XMLDoc.loadXML(ServerAnswer)
					resultStr = resultStr & "<p>" & XMLDoc.SelectNodes("//PAGE").length & " " & dlgInstances & "</p>"
					
				else
					resultStr = resultStr & "<p><b>" & dlgError & ": " & dlgContentClassNotFound & "!</b></p>"
				end if
				
			else
				resultStr = resultStr & "<p><b>" & dlgError & ": " & dlgFolderNotFound & "!</b></p>"
			end if
			
		else
			resultStr = resultStr & "<p><b>" & dlgError & ": " & dlgNoRights & "!</b></p>"
		end if		

	next

	'Logout RQLAdmin
	xmlSendDoc=	"<IODATA loginguid=""" & RqlAdmLoginGUID & """>"&_
					"<ADMINISTRATION>"&_
						"<LOGOUT guid=""" & RqlAdmLoginGUID & """ />"&_
					"</ADMINISTRATION>"&_
				"</IODATA>"
	ServerAnswer = objIO.ServerExecuteXml (xmlSendDoc, sError)

else
	resultStr = "<p><b>" & dlgError & "!</b></p><p>" & dlgInsufficientRights & "!</p>"
end if

Set ProjekteList = nothing

' säubern
set XMLDoc = nothing
set XMLFoldersDoc = nothing
set XMLProjDoc = nothing
set XMLTemplatesDoc = nothing
set objIO = nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" type="text/css" href="../../stylesheets/ioStyleSheet.css" />
<title><%=pluginTitle %></title>
</head>
<body link="navy" alink="navy" vlink="navy" bgcolor="#ffffff" background="../../icons/back5.gif">
<form name="TemplateDependenceCheckerForm" method="post" action="#">
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
		<td align="left" valign="top" class="label" width="100%"><%=dlgHeadline %>:</td>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		</tr>
		<tr>
		<td height="5" colspan="3"></td>
		</tr>
		<tr>
		<td></td>
		<td><%=resultStr %></td>
		<td></td>
		</tr>
		<tr>
		<td height="20" colspan="3"><img src="../../icons/transparent.gif" width="1" height="20" border="0" alt=""></td>
		</tr>
		<tr>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		<td align="right" valign="top"><button id="btn2" type="button" onclick="self.close()"><%=dlgClose %></button></td>
		<td width="25"><img src="../../icons/transparent.gif" width="25" height="1" border="0" alt=""></td>
		</tr>
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