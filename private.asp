<%@ Language="VBScript" %>
<!-- #include file="common.inc" -->
<%
' Looking for URL params
Catalog = Request.QueryString("catalog")

if Len(Catalog) > 0 then
	' Init
	CatalogPath = Server.MapPath(CatalogDir & WebDirDelimiter & Catalog)

	' If catalog is OK
	if FileSystem.FolderExists(CatalogPath) then
		' Getting environment
		CatalogConfig = GetFileContent(CatalogPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf
		DefaultCatalogConfig = GetFileContent(RootPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf

		CatalogLanguage = GetConfigValue(CatalogConfig, CatalogLanguageKey, GetConfigValue(DefaultCatalogConfig, CatalogLanguageKey, DefaultCatalogLanguage))
		CatalogStyle = GetConfigValue(CatalogConfig, CatalogStyleKey, GetConfigValue(DefaultCatalogConfig, CatalogStyleKey, DefaultCatalogStyle))
		CatalogTitle = GetConfigValue(CatalogConfig, CatalogTitleKey, GetConfigValue(DefaultCatalogConfig, CatalogTitleKey, Catalog))

		LoadLanguage CatalogLanguage

		' Start of page
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=" & Quote & "Content-Type" & Quote & " content=" & Quote & "text/html; charset=" & LanguageCharset & Quote & "/>"
		if Len(CatalogStyle) > 0 then
			Response.Write "<link rel=" & Quote & "stylesheet" & Quote & " type=" & Quote & "text/css" & Quote & " href=" & Quote & StyleDir & WebDirDelimiter & CatalogStyle & ".css" & Quote & "/>"
		end if
		Response.Write "<link rel=" & Quote & "icon" & Quote & " type=" & Quote & "image/png" & Quote & " href=" & Quote & "favicon.png" & Quote & "/>"
		Response.Write "<title>" & CatalogTitle & "</title>"
		Response.Write "</head>"
		Response.Write "<body>"
		Response.Write "<table width=" & Quote & "100%" & Quote & " height=" & Quote & "100%" & Quote & ">"
		Response.Write "<tbody>"
		Response.Write "<tr valign=" & Quote & "center" & Quote & ">"
		Response.Write "<td align=" & Quote & "center" & Quote & ">"

		' Form
		Response.Write "<form method=" & Quote & "get" & Quote & " action=" & Quote & "album.asp" & Quote & ">"
		Response.Write "<p><a href=" & Quote & "catalog.asp?name=" & Catalog & Quote & " title=" & Quote & LanguageMainPageTitle & Quote & ">" & LanguageMainPage & "</a></p>"
		Response.Write "<p>" & LanguagePrivateText & "</p>"
		Response.Write "<p><input type=" & Quote & "hidden" & Quote & " name=" & Quote & "catalog" & Quote & " value=" & Quote & Catalog & Quote & "></p>"
		Response.Write "<p><input type=" & Quote & "text" & Quote & " name=" & Quote & "name" & Quote & " size=" & Quote & "64" & Quote & "></p>"
		Response.Write "<p><input type=" & Quote & "submit" & Quote & " value=" & Quote & LanguagePrivateAction & Quote & "></p>"
		Response.Write "</form>"

		' End of page
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
	end if
end if
%>
