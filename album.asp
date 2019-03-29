<%@ Language="VBScript" %>
<!-- #include file="utils.inc" -->
<!-- #include file="imaging.inc" -->
<!-- #include file="security.inc" -->
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
    Response.Write "<meta http-equiv=" & Quote & "Content-Type" & Quote & " content=" & Quote & "text/html; charset=" & LanguageCharset & Quote & ">"

    if Len(CatalogStyle) > 0 then
      Response.Write "<link rel=" & Quote & "stylesheet" & Quote & " type=" & Quote & "text/css" & Quote & " href=" & Quote & StyleDir & WebDirDelimiter & CatalogStyle & WebDirDelimiter & "album.css" & Quote & ">"
    end if

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
    Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><br></td></tr>"
    Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & " align=" & Quote & "center" & Quote & ">" & LanguageThumbnailHint & "</td></tr>"
    Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><br></td></tr>"
    Response.Write "<tr><td colspan=" & Quote & AlbumThumbnailGridCols & Quote & "><hr></td></tr>"

    ' Looking for images
    FirstImageNum = (PageNum - 1) * AlbumThumbnailGridCols * AlbumThumbnailGridRows + 1
    LastImageNum = PageNum * AlbumThumbnailGridCols * AlbumThumbnailGridRows
    ImageCounter = 0
    OnPageImageCount = 0

    for each File in FileSystem.GetFolder(AlbumPath).Files
      Extension = "." & FileSystem.GetExtensionName(File.Name)

      if ContainsValue(ImageExtensions, Lcase(Extension)) and File.Name <> AlbumLogoFileName and Left(Right(Lcase(File.Name), Len(ThumbnailPostfix) + Len(Extension)), Len(ThumbnailPostfix)) <> ThumbnailPostfix then
        ImageCounter = ImageCounter + 1
        ImageFileName = File.Name
        ThumbnailFileName = Left(ImageFileName, Len(ImageFileName) - Len(Extension)) & ThumbnailPostfix & Extension

        if ImageCounter >= FirstImageNum and ImageCounter <= LastImageNum then
          ' Checking thumbnail
          if not FileSystem.FileExists(AlbumPath & OSDirDelimiter & ThumbnailFileName) then
            GenerateImagePreview AlbumPath & OSDirDelimiter & ImageFileName, AlbumPath & OSDirDelimiter & ThumbnailFileName, AlbumThumbnailSize, False
            CopyNTFSSecuritySettings AlbumPath & OSDirDelimiter & ImageFileName, AlbumPath & OSDirDelimiter & ThumbnailFileName
          end if

          if (ImageCounter - 1) mod AlbumThumbnailGridCols = 0 then Response.Write "<tr valign=" & Quote & "center" & Quote & ">"

          ImageLink = "image.asp?catalog=" & Catalog & "&album=" & Album & "&name=" & ImageFileName
          Response.Write "<td align=" & Quote & "center" & Quote & " width=" & Quote & AlbumCellSize & Quote & " height=" & Quote & AlbumCellSize & Quote & "><a href=" & Quote & ImageLink & Quote & " title=" & Quote & ImageFileName & Quote & "><img src=" & Quote & CatalogDir & WebDirDelimiter & Catalog & WebDirDelimiter & Album & WebDirDelimiter & ThumbnailFileName & Quote & " alt=" & Quote & ImageFileName & Quote & "></a></td>"
          OnPageImageCount = OnPageImageCount + 1

          if ImageCounter mod AlbumThumbnailGridCols = 0 then Response.Write "</tr>"
        end if
      end if
    next

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
