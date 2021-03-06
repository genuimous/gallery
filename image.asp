<%@ Language="VBScript" %>
<!-- #include file="common.inc" -->
<!-- #include file="imaging.inc" -->
<%
' Looking for URL params
Catalog = Request.QueryString("catalog")
Album = Request.QueryString("album")
ImageFileName = Request.QueryString("name")

if Len(Catalog) > 0 and Len(Album) > 0 and Len(ImageFileName) > 0 then
	' Init
	CatalogPath = Server.MapPath(CatalogDir & WebDirDelimiter & Catalog)
	AlbumPath = Server.MapPath(CatalogDir & WebDirDelimiter & Catalog & WebDirDelimiter & Album)

	' If album is OK
	if FileSystem.FileExists(AlbumPath & OSDirDelimiter & ImageFileName) then
		' Getting environment
		CatalogConfig = GetFileContent(CatalogPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf
		DefaultCatalogConfig = GetFileContent(RootPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf

		CatalogLanguage = GetConfigValue(CatalogConfig, CatalogLanguageKey, GetConfigValue(DefaultCatalogConfig, CatalogLanguageKey, ""))
		CatalogStyle = GetConfigValue(CatalogConfig, CatalogStyleKey, GetConfigValue(DefaultCatalogConfig, CatalogStyleKey, ""))

		AlbumConfig = GetFileContent(AlbumPath & OSDirDelimiter & AlbumConfigFileName) & vbCrLf
		DefaultAlbumConfig = GetFileContent(RootPath & OSDirDelimiter & AlbumConfigFileName) & vbCrLf

		AlbumLogoFileName = GetConfigValue(AlbumConfig, AlbumLogoFileNameKey, GetConfigValue(DefaultAlbumConfig, AlbumLogoFileNameKey, DefaultAlbumLogoFileName))
		AlbumThumbnailGridCols = GetConfigValue(AlbumConfig, AlbumThumbnailGridColsKey, GetConfigValue(DefaultAlbumConfig, AlbumThumbnailGridColsKey, DefaultAlbumThumbnailGridCols))
		AlbumThumbnailGridRows = GetConfigValue(AlbumConfig, AlbumThumbnailGridRowsKey, GetConfigValue(DefaultAlbumConfig, AlbumThumbnailGridRowsKey, DefaultAlbumThumbnailGridRows))

		ImageConfig = GetFileContent(AlbumPath & OSDirDelimiter & ImageConfigFileName) & vbCrLf
		DefaultImageConfig = GetFileContent(RootPath & OSDirDelimiter & ImageConfigFileName) & vbCrLf

		ImageNavigation = GetConfigValue(ImageConfig, ImageNavigationKey, GetConfigValue(DefaultImageConfig, ImageNavigationKey, DefaultImageNavigation))
		ImagePreviewSize = GetConfigValue(ImageConfig, ImagePreviewSizeKey, GetConfigValue(DefaultImageConfig, ImagePreviewSizeKey, DefaultImagePreviewSize))

		LoadLanguage CatalogLanguage

		' Start of page
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=" & Quote & "Content-Type" & Quote & " content=" & Quote & "text/html; charset=" & LanguageCharset & Quote & "/>"
		if Len(CatalogStyle) > 0 then
			Response.Write "<link rel=" & Quote & "stylesheet" & Quote & " type=" & Quote & "text/css" & Quote & " href=" & Quote & StyleDir & WebDirDelimiter & CatalogStyle & ".css" & Quote & "/>"
		end if
		Response.Write "<link rel=" & Quote & "icon" & Quote & " type=" & Quote & "image/png" & Quote & " href=" & Quote & "favicon.png" & Quote & "/>"
		Response.Write "<title>" & ImageFileName & "</title>"
		Response.Write "</head>"
		Response.Write "<body>"
		Response.Write "<table width=" & Quote & "100%" & Quote & " height=" & Quote & "100%" & Quote & ">"
		Response.Write "<tbody>"
		Response.Write "<tr valign=" & Quote & "center" & Quote & ">"
		Response.Write "<td align=" & Quote & "center" & Quote & ">"
		Response.Write "<table align=" & Quote & "center" & Quote & ">"
		Response.Write "<tbody>"
		Response.Write "<tr><td colspan=" & Quote & "3" & Quote & "><br></td></tr>"

		' Looking for images
		set FileList = FileSystem.GetFolder(AlbumPath).Files

		ImageCounter = 0
		ImageFileNum = 0
		PreviousImageFileName = ""
		NextImageFileName = ""
		TotalImageCount = 0

		for each File in FileList
			Extension = "." & FileSystem.GetExtensionName(File.Name)

			if ContainsValue(ImageExtensions, Extension) and File.Name <> AlbumLogoFileName and Left(Right(Lcase(File.Name), Len(ThumbnailPostfix) + Len(Extension)), Len(ThumbnailPostfix)) <> ThumbnailPostfix and Left(Right(Lcase(File.Name), Len(PreviewPostfix) + Len(Extension)), Len(PreviewPostfix)) <> PreviewPostfix then
				ImageCounter = ImageCounter + 1
				TotalImageCount = TotalImageCount + 1

				if ImageFileNum = 0 then
					if File.Name = ImageFileName then
						ImageFileNum = ImageCounter
					else
						PreviousImageFileName = File.Name
					end if
				else
					if Len(NextImageFileName) = 0 then
						NextImageFileName = File.Name
					end if
				end if
			end if
		next

		set FileList = Nothing

		Extension = "." & FileSystem.GetExtensionName(ImageFileName)

		' Image
		if ContainsValue(ImageExtensions, Extension) then
			PreviewFileName = Left(ImageFileName, Len(ImageFileName) - Len(Extension)) & PreviewPostfix & Extension

			' Checking preview
			if not FileSystem.FileExists(AlbumPath & OSDirDelimiter & PreviewFileName) then
				GenerateImagePreview AlbumPath & OSDirDelimiter & ImageFileName, AlbumPath & OSDirDelimiter & PreviewFileName, ImagePreviewSize, False
			end if

			Response.Write "<tr><td colspan=" & Quote & "3" & Quote & " align=" & Quote & "center" & Quote & ">" & LanguagePreviewHint & " </td></tr>"
			Response.Write "<tr><td colspan=" & Quote & "3" & Quote & "><br></td></tr>"
			Response.Write "<tr><td colspan=" & Quote & "2" & Quote & " align=" & Quote & "left" & Quote & ">" & ImageFileName & "</td><td align=" & Quote & "right" & Quote & ">" & ImageFileNum & "/" & TotalImageCount & "</td></tr>"
			Response.Write "<tr><td colspan=" & Quote & "3" & Quote & "><br></td></tr>"
			Response.Write "<tr><td colspan=" & Quote & "3" & Quote & "><a download href=" & Quote & CatalogDir & WebDirDelimiter & Catalog & WebDirDelimiter & Album & WebDirDelimiter & ImageFileName & Quote & " title=" & Quote & ImageFileName & Quote & "><img src=" & Quote & CatalogDir & WebDirDelimiter & Catalog & WebDirDelimiter & Album & WebDirDelimiter & PreviewFileName & Quote & " alt=" & Quote & ImageFileName & Quote & "/></a></td></tr>"
			Response.Write "<tr><td colspan=" & Quote & "3" & Quote & "><br></td></tr>"
		end if

		' Navigator
		if ImageNavigation then
			PageNum = -Int(-ImageFileNum / (AlbumThumbnailGridCols * AlbumThumbnailGridRows))

			if Len(PreviousImageFileName) > 0 then
				PreviousImageLink = "image.asp?catalog=" & Catalog & "&album=" & Album & "&name=" & PreviousImageFileName
			else
				PreviousImageLink = ""
			end if
			AlbumLink = "album.asp?catalog=" & Catalog & "&name=" & Album & "&page=" & PageNum
			if Len(NextImageFileName) > 0 then
				NextImageLink = "image.asp?catalog=" & Catalog & "&album=" & Album & "&name=" & NextImageFileName
			else
				NextImageLink = ""
			end if

			Response.Write "<tr>"
			Response.Write "<td width=" & Quote & "200" & Quote & " align=" & Quote & "left" & Quote & ">"
			if Len(PreviousImageLink) > 0 then
				Response.Write "<a href=" & Quote & PreviousImageLink & Quote & " title=" & Quote & LanguagePreviousTitle & Quote & ">&lt;&lt; " & LanguagePrevious &"</a>"
			end if
			Response.Write "</td>"
			Response.Write "<td align=" & Quote & "center" & Quote & "><a href=" & Quote & AlbumLink & Quote & " title=" & Quote & LanguageIndexTitle & Quote & ">" & LanguageIndex &"</a></td>"
			Response.Write "<td width=" & Quote & "200" & Quote & " align=" & Quote & "right" & Quote & ">"
			if Len(NextImageLink) > 0 then
				Response.Write "<a href=" & Quote & NextImageLink & Quote & " title=" & Quote & LanguageNextTitle & Quote & ">" & LanguageNext &" &gt;&gt;</a>"
			end if
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "<tr><td colspan=" & Quote & "3" & Quote & "><br></td></tr>"
		end if

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
