<%@ Language="VBScript" %>
<!-- #include file="common.inc" -->
<!-- #include file="imaging.inc" -->
<%
' Looking for URL params
Catalog = Request.QueryString("name")

if Len(Catalog) > 0 then
	' Init
	CatalogPath = Server.MapPath(CatalogDir & WebDirDelimiter & Catalog)

	' If catalog is OK
	if FileSystem.FolderExists(CatalogPath) then
		' Getting environment
		CatalogConfig = GetFileContent(CatalogPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf
		DefaultCatalogConfig = GetFileContent(RootPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf

		CatalogInfo = ClearText(GetFileContent(CatalogPath & OSDirDelimiter & CatalogInfoFileName))

		CatalogDefense = GetConfigValue(CatalogConfig, CatalogDefenseKey, GetConfigValue(DefaultCatalogConfig, CatalogDefenseKey, DefaultCatalogDefense))
		CatalogLanguage = GetConfigValue(CatalogConfig, CatalogLanguageKey, GetConfigValue(DefaultCatalogConfig, CatalogLanguageKey, DefaultCatalogLanguage))
		CatalogStyle = GetConfigValue(CatalogConfig, CatalogStyleKey, GetConfigValue(DefaultCatalogConfig, CatalogStyleKey, DefaultCatalogStyle))
		CatalogTitle = GetConfigValue(CatalogConfig, CatalogTitleKey, GetConfigValue(DefaultCatalogConfig, CatalogTitleKey, Catalog))
		CatalogLogoSize = GetConfigValue(CatalogConfig, CatalogLogoSizeKey, GetConfigValue(DefaultCatalogConfig, CatalogLogoSizeKey, DefaultCatalogLogoSize))
		CatalogAuthor = GetConfigValue(CatalogConfig, CatalogAuthorKey, GetConfigValue(DefaultCatalogConfig, CatalogAuthorKey, Catalog))
		CatalogAuthorSite = GetConfigValue(CatalogConfig, CatalogAuthorSiteKey, GetConfigValue(DefaultCatalogConfig, CatalogAuthorSiteKey, ""))
		CatalogAuthorEmail = GetConfigValue(CatalogConfig, CatalogAuthorEmailKey, GetConfigValue(DefaultCatalogConfig, CatalogAuthorEmailKey, ""))
		CatalogAuthorPhone = GetConfigValue(CatalogConfig, CatalogAuthorPhoneKey, GetConfigValue(DefaultCatalogConfig, CatalogAuthorPhoneKey, ""))

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
		Response.Write "<table align=" & Quote & "center" & Quote & ">"
		Response.Write "<tbody>"
		Response.Write "<tr><td><br></tr></td>"

		' Header
		Response.Write "<tr><td align=" & Quote & "center" & Quote & "><h1>" & CatalogTitle & "</h1></td></tr>"
		Response.Write "<tr><td><br></tr></td>"

		' Info
		if Len(CatalogInfo) > 0 then
			Response.Write "<tr><td>" & CatalogInfo & "</tr></td>"
			Response.Write "<tr><td><br></tr></td>"
		end if

		' Albums
		if not CatalogDefense then
			Response.Write "<tr><td><hr></tr></td>"
			Response.Write "<tr><td>"
			Response.Write "<table align=" & Quote & "center" & Quote & "cellpadding=" & Quote & "5" & Quote & ">"
			Response.Write "<tbody>"

			' Looking for albums
			set FolderList = FileSystem.GetFolder(CatalogPath).SubFolders

			for each Folder in FolderList
				if Left(Folder.Name, 1) <> "." then
					Album = Folder.Name
					AlbumPath = Server.MapPath(CatalogDir & WebDirDelimiter & Catalog & WebDirDelimiter & Album)
					AlbumConfig = GetFileContent(AlbumPath & OSDirDelimiter & AlbumConfigFileName) & vbCrLf
					DefaultAlbumConfig = GetFileContent(RootPath & OSDirDelimiter & AlbumConfigFileName) & vbCrLf
					AlbumLink = "album.asp?catalog=" & Catalog & "&name=" & Album

					AlbumLogoFileName = GetConfigValue(AlbumConfig, AlbumLogoFileNameKey, GetConfigValue(DefaultAlbumConfig, AlbumLogoFileNameKey, DefaultAlbumLogoFileName))
					AlbumTitle = GetConfigValue(AlbumConfig, AlbumTitleKey, GetConfigValue(DefaultAlbumConfig, AlbumTitleKey, Album))
					AlbumDescription = GetConfigValue(AlbumConfig, AlbumDescriptionKey, GetConfigValue(DefaultAlbumConfig, AlbumDescriptionKey, ""))

					' Checking logo
					if not FileSystem.FileExists(AlbumPath & OSDirDelimiter & AlbumLogoFileName) then
						' Looking for suitable image
						ImageFileName = ""

						set FileList = FileSystem.GetFolder(AlbumPath).Files

						for each File in FileList
							Extension = "." & FileSystem.GetExtensionName(File.Name)

							if ContainsValue(ImageExtensions, Extension) and Left(Right(Lcase(File.Name), Len(ThumbnailPostfix) + Len(Extension)), Len(ThumbnailPostfix)) <> ThumbnailPostfix and Left(Right(Lcase(File.Name), Len(PreviewPostfix) + Len(Extension)), Len(PreviewPostfix)) <> PreviewPostfix then
								ImageFileName = File.Name
								exit for
							end if
						next

						set FileList = Nothing

						if Len(ImageFileName) > 0 then
							GenerateImagePreview AlbumPath & OSDirDelimiter & ImageFileName, AlbumPath & OSDirDelimiter & AlbumLogoFileName, CatalogLogoSize, True
						end if
					end if

					Response.Write "<tr>"
					Response.Write "<td><a href=" & Quote & AlbumLink & Quote & " title=" & Quote & AlbumTitle & Quote & "><img width=" & Quote & CatalogLogoSize & Quote & " height=" & Quote & CatalogLogoSize & Quote & " src=" & Quote & CatalogDir & "/" & Catalog & "/"& Album & "/" & AlbumLogoFileName & Quote & "></a></td>"
					Response.Write "<td width=" & Quote & "100%" & Quote & " align=" & Quote & "left" & Quote & " valign=" & Quote & "top" & Quote & "><h2><a href=" & Quote & AlbumLink & Quote & " title=" & Quote & AlbumTitle & Quote & ">" & AlbumTitle & "</a></h2>" & AlbumDescription & "</td>"
					Response.Write "</tr>"
				end if
			next

			set FolderList = Nothing

			Response.Write "</tbody>"
			Response.Write "</table>"
			Response.Write "</tr></td>"
			Response.Write "<tr><td><hr></tr></td>"
			Response.Write "<tr><td><br></tr></td>"
		end if

		' Footer
		Response.Write "<tr><td align=" & Quote & "center" & Quote & ">"
		Response.Write Copyright & " "
		if CalalogYear < CurrentYear then
			Response.Write CalalogYear & "-" & CurrentYear
		else
			Response.Write CurrentYear
		end if
		Response.Write " "
		if Len(CatalogAuthorSite) > 0 then
			Response.Write "<a href=" & Quote & CatalogAuthorSite & Quote & ">" & CatalogAuthor & "</a>"
		else
			Response.Write CatalogAuthor
		end if
		if Len(CatalogAuthorEmail) > 0 then
			Response.Write " " & LanguageEMail & ": " & "<a href=" & Quote & "mailto:" & CatalogAuthorEmail & Quote & ">" & CatalogAuthorEmail & "</a>"
		end if
		if Len(CatalogAuthorPhone) > 0 then
			Response.Write " " & LanguageTelephone & ": " & CatalogAuthorPhone
		end if
		Response.Write "</tr></td>"
		Response.Write "<tr><td><br></tr></td>"

		' End of page
		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</tbody>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
	end if
end if
%>
