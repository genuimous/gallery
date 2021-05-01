<%@ Language="VBScript" %>
<!-- #include file="common.inc" -->
<!-- #include file="utils.inc" -->
<!-- #include file="math.inc" -->
<!-- #include file="imaging.inc" -->
<%
' Looking at URL params
Catalog = Request.QueryString("catalog")
Album = Request.QueryString("name")
if Len(Request.QueryString("page")) > 0 then
	if IsNumeric(Request.QueryString("page")) then
		PageNum = CLng(Request.QueryString("page"))
	else
		PageNum = 0
	end if
else
	PageNum = 1
end if

if Len(Catalog) > 0 and Len(Album) > 0 and PageNum >= 1 then
	' Init
	CatalogPath = Server.MapPath(CatalogDir & WebDirDelimiter & Catalog)
	AlbumPath = Server.MapPath(CatalogDir & WebDirDelimiter & Catalog & WebDirDelimiter & Album)

	' If album is OK
	if FileSystem.FolderExists(AlbumPath) then
		' Getting environment
		CatalogConfig = GetFileContent(CatalogPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf
		DefaultCatalogConfig = GetFileContent(RootPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf

		CatalogLanguage = GetConfigValue(CatalogConfig, CatalogLanguageKey, GetConfigValue(DefaultCatalogConfig, CatalogLanguageKey, DefaultCatalogLanguage))
		CatalogStyle = GetConfigValue(CatalogConfig, CatalogStyleKey, GetConfigValue(DefaultCatalogConfig, CatalogStyleKey, DefaultCatalogStyle))

		AlbumConfig = GetFileContent(AlbumPath & OSDirDelimiter & AlbumConfigFileName) & vbCrLf
		DefaultAlbumConfig = GetFileContent(RootPath & OSDirDelimiter & AlbumConfigFileName) & vbCrLf

		AlbumLogoFileName = GetConfigValue(AlbumConfig, AlbumLogoFileNameKey, GetConfigValue(DefaultAlbumConfig, AlbumLogoFileNameKey, DefaultAlbumLogoFileName))
		AlbumArchiveFileName = GetConfigValue(AlbumConfig, AlbumArchiveFileNameKey, GetConfigValue(DefaultAlbumConfig, AlbumArchiveFileNameKey, DefaultArchiveFileNameKey))
		AlbumTitle = GetConfigValue(AlbumConfig, AlbumTitleKey, GetConfigValue(DefaultAlbumConfig, AlbumTitleKey, Album))
		AlbumDescription = GetConfigValue(AlbumConfig, AlbumDescriptionKey, GetConfigValue(DefaultAlbumConfig, AlbumDescriptionKey, ""))
		AlbumThumbnailSize = GetConfigValue(AlbumConfig, AlbumThumbnailSizeKey, GetConfigValue(DefaultAlbumConfig, AlbumThumbnailSizeKey, DefaultAlbumThumbnailSize))
		AlbumThumbnailGridCols = GetConfigValue(AlbumConfig, AlbumThumbnailGridColsKey, GetConfigValue(DefaultAlbumConfig, AlbumThumbnailGridColsKey, DefaultAlbumThumbnailGridCols))
		AlbumThumbnailGridRows = GetConfigValue(AlbumConfig, AlbumThumbnailGridRowsKey, GetConfigValue(DefaultAlbumConfig, AlbumThumbnailGridRowsKey, DefaultAlbumThumbnailGridRows))
		AlbumThumbnailIndent = GetConfigValue(AlbumConfig, AlbumThumbnailIndentKey, GetConfigValue(DefaultAlbumConfig, AlbumThumbnailIndentKey, DefaultAlbumThumbnailIndent))

		AlbumCellSize = AlbumThumbnailSize + AlbumThumbnailIndent * 2

		LoadLanguage CatalogLanguage

		' Start of page
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=" & Quote & "Content-Type" & Quote & " content=" & Quote & "text/html; charset=" & LanguageCharset & Quote & "/>"
		if Len(CatalogStyle) > 0 then
			Response.Write "<link rel=" & Quote & "stylesheet" & Quote & " type=" & Quote & "text/css" & Quote & " href=" & Quote & StyleDir & WebDirDelimiter & CatalogStyle & WebDirDelimiter & "album.css" & Quote & "/>"
		end if
		Response.Write "<link rel=" & Quote & "icon" & Quote & " type=" & Quote & "image/png" & Quote & " href=" & Quote & "favicon.png" & Quote & "/>"
		Response.Write "<title>" & AlbumTitle & "</title>"
		Response.Write "</head>"
		Response.Write "<body>"
		Response.Write "<table width=" & Quote & "100%" & Quote & " height=" & Quote & "100%" & Quote & ">"
		Response.Write "<tbody>"
		Response.Write "<tr valign=" & Quote & "center" & Quote & ">"
		Response.Write "<td align=" & Quote & "center" & Quote & ">"
		Response.Write "<table align=" & Quote & "center" & Quote & " width=" & Quote & AlbumThumbnailGridCols * AlbumCellSize & Quote & ">"
		Response.Write "<tbody>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><br></td></tr>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & " align=" & Quote & "center" & Quote & "><h1>" & AlbumTitle & "</h1></td></tr>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & " align=" & Quote & "center" & Quote & "><h2>" & AlbumDescription & "</h2></td></tr>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><br></td></tr>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & " align=" & Quote & "center" & Quote & "><a href=" & Quote & "catalog.asp?name=" & Catalog & Quote & " title=" & Quote & LanguageMainPageTitle & Quote & ">" & LanguageMainPage & "</a></td></tr>"
		if Len(AlbumArchiveFileName) > 0 then
				if FileSystem.FileExists(AlbumPath & OSDirDelimiter & AlbumArchiveFileName) then
						Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & " align=" & Quote & "center" & Quote & "><a download href=" & Quote & CatalogDir & WebDirDelimiter & Catalog & WebDirDelimiter & Album & WebDirDelimiter & AlbumArchiveFileName & Quote & " title=" & Quote & LanguageDownloadAllTitle & Quote & ">" & LanguageDownloadAll & "</a></td></tr>"
				end if
		end if
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><br></td></tr>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & " align=" & Quote & "center" & Quote & ">" & LanguageThumbnailHint & "</td></tr>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><br></td></tr>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><hr></td></tr>"

		' Looking for images
		set FileList = FileSystem.GetFolder(AlbumPath).Files

		FirstImageNum = (PageNum - 1) * AlbumThumbnailGridCols * AlbumThumbnailGridRows + 1
		LastImageNum = PageNum * AlbumThumbnailGridCols * AlbumThumbnailGridRows
		ImageCounter = 0
		OnPageImageCount = 0

		for each File in FileList
			Extension = "." & FileSystem.GetExtensionName(File.Name)

			if ContainsValue(ImageExtensions, Extension) and File.Name <> AlbumLogoFileName and Left(Right(Lcase(File.Name), Len(ThumbnailPostfix) + Len(Extension)), Len(ThumbnailPostfix)) <> ThumbnailPostfix and Left(Right(Lcase(File.Name), Len(PreviewPostfix) + Len(Extension)), Len(PreviewPostfix)) <> PreviewPostfix then
				ImageCounter = ImageCounter + 1
				ImageFileName = File.Name
				ThumbnailFileName = Left(File.Name, Len(File.Name) - Len(Extension)) & ThumbnailPostfix & Extension

				if ImageCounter >= FirstImageNum and ImageCounter <= LastImageNum then
					' Checking thumbnail
					if not FileSystem.FileExists(AlbumPath & OSDirDelimiter & ThumbnailFileName) then
						GenerateImagePreview AlbumPath & OSDirDelimiter & ImageFileName, AlbumPath & OSDirDelimiter & ThumbnailFileName, AlbumThumbnailSize, False
					end if

					if (ImageCounter - 1) mod AlbumThumbnailGridCols = 0 then Response.Write "<tr valign=" & Quote & "center" & Quote & ">"

					ImageLink = "image.asp?catalog=" & Catalog & "&album=" & Album & "&name=" & ImageFileName
					Response.Write "<td align=" & Quote & "center" & Quote & " width=" & Quote & AlbumCellSize & Quote & " height=" & Quote & AlbumCellSize & Quote & "><a href=" & Quote & ImageLink & Quote & " title=" & Quote & ImageFileName & Quote & "><img src=" & Quote & CatalogDir & WebDirDelimiter & Catalog & WebDirDelimiter & Album & WebDirDelimiter & ThumbnailFileName & Quote & " alt=" & Quote & ImageFileName & Quote & "/></a></td>"
					OnPageImageCount = OnPageImageCount + 1

					if ImageCounter mod AlbumThumbnailGridCols = 0 then Response.Write "</tr>"
				end if
			end if
		next

		set FileList = Nothing

		TotalImageCount = ImageCounter

		if OnPageImageCount mod AlbumThumbnailGridCols > 0 then
			for MissedImageCounter = 1 to AlbumThumbnailGridCols - (OnPageImageCount mod AlbumThumbnailGridCols)
				Response.Write "<td align=" & Quote & "center" & Quote & " width=" & Quote & AlbumCellSize & Quote & " height=" & Quote & AlbumCellSize & Quote & "></td>"
			next

			Response.Write "</tr>"
		end if

		' Page list
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><hr></td></tr>"
		Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><br></td></tr>"

		TotalPageCount = RoundUp(TotalImageCount / (AlbumThumbnailGridCols * AlbumThumbnailGridRows))

		if TotalPageCount > 1 then
			Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & " align=" & Quote & "center" & Quote & ">"

			for PageCounter = 1 to TotalPageCount
				if PageCounter > 1 then Response.Write " "

				if PageCounter = PageNum then
					Response.Write "[" & PageCounter & "]"
				else
					Response.Write "<a href=" & Quote & "album.asp?catalog=" & Catalog & "&name=" & Album & "&page=" & PageCounter & Quote & " title=" & Quote & LanguageIndexNumTitle & PageCounter & Quote & ">[" & PageCounter & "]</a>"
				end if
			next

			Response.Write "</td></tr>"
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
