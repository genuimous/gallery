<%@ Language="VBScript" %>
<!-- #include file="utils.inc" -->
<!-- #include file="imaging.inc" -->
<!-- #include file="security.inc" -->
<%
' Looking at URL params
Catalog = Request.QueryString("name")

if Len(Catalog) > 0 then
  ' Init
  CatalogPath = Server.MapPath(CatalogDir & WebDirDelimiter & Catalog)

  ' If catalog is OK
  if FileSystem.FolderExists(CatalogPath) then
    ' Getting environment
    CatalogConfig = GetFileContent(CatalogPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf
    DefaultCatalogConfig = GetFileContent(RootPath & OSDirDelimiter & CatalogConfigFileName) & vbCrLf

    CatalogIntroText = GetFileContent(CatalogPath & OSDirDelimiter & CatalogIntroTextFileName)
    CatalogOutroText = GetFileContent(CatalogPath & OSDirDelimiter & CatalogOutroTextFileName)

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
    Response.Write "<meta http-equiv=" & Quote & "Content-Type" & Quote & " content=" & Quote & "text/html; charset=" & LanguageCharset & Quote & ">"

    if Len(CatalogStyle) > 0 then
      Response.Write "<link rel=" & Quote & "stylesheet" & Quote & " type=" & Quote & "text/css" & Quote & " href=" & Quote & StyleDir & WebDirDelimiter & CatalogStyle & WebDirDelimiter & "catalog.css" & Quote & ">"
    end if

    Response.Write "<title>" & CatalogTitle & "</title>"
    Response.Write "</head>"
    Response.Write "<body>"
    Response.Write "<center>"
    Response.Write "<div id=" & Quote & "title" & Quote & "><h1>" & CatalogTitle & "</h1></div>"
    Response.Write "<div id=" & Quote & "content" & Quote & ">"
    Response.Write "<p>" & CatalogIntroText& "</p>"
    Response.Write "<hr>"
    Response.Write "<table align=" & Quote & "center" & Quote & ">"
    Response.Write "<tbody>"

    ' Looking for albums
    dim Albums()
    AlbumCount = 0
    for each Folder in FileSystem.GetFolder(CatalogPath).SubFolders
      AlbumCount = AlbumCount + 1
      redim preserve Albums(AlbumCount - 1)
      Albums(AlbumCount - 1) = Folder.Name
    next

    if AlbumCount > 1 then
      for AlbumCounter = 0 to (AlbumCount - 1) \ 2
        Album = Albums(AlbumCounter)
        Albums(AlbumCounter) = Albums(AlbumCount - AlbumCounter - 1)
        Albums(AlbumCount - AlbumCounter - 1) = Album
      next
    end if

    for each Album in Albums
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

        for each File in FileSystem.GetFolder(AlbumPath).Files
          Extension = "." & Lcase(FileSystem.GetExtensionName(File.Name))

          if ContainsValue(ImageExtensions, Extension) and Left(Right(Lcase(File.Name), Len(ThumbnailPostfix) + Len(Extension)), Len(ThumbnailPostfix)) <> ThumbnailPostfix then
            ImageFileName = File.Name
            exit for
          end if
        next

        if Len(ImageFileName) > 0 then
          GenerateImagePreview AlbumPath & OSDirDelimiter & ImageFileName, AlbumPath & OSDirDelimiter & AlbumLogoFileName, CatalogLogoSize, True
          CopyNTFSSecuritySettings AlbumPath & OSDirDelimiter & ImageFileName, AlbumPath & OSDirDelimiter & AlbumLogoFileName
        end if
      end if

      Response.Write "<tr>"
      Response.Write "<td><a href=" & Quote & AlbumLink & Quote & " title=" & Quote & AlbumTitle & Quote & "><img width=" & Quote & CatalogLogoSize & Quote & " height=" & Quote & CatalogLogoSize & Quote & " src=" & Quote & CatalogDir & "/" & Catalog & "/"& Album & "/" & AlbumLogoFileName & Quote & "></a></td>"
      Response.Write "<td width=" & Quote & "100%" & Quote & " align=" & Quote & "left" & Quote & " valign=" & Quote & "top" & Quote & "><h2><a href=" & Quote & AlbumLink & Quote & " title=" & Quote & AlbumTitle & Quote & ">" & AlbumTitle & "</a></h2>" & AlbumDescription & "</td>"
      Response.Write "</tr>"
    next

    Response.Write "</tbody>"
    Response.Write "</table>"
    Response.Write "<hr>"
    Response.Write "<p>" & CatalogOutroText & "</p>"
    Response.Write "</div>"

    ' Copyright
    CalalogYear = DatePart("yyyy", FileSystem.GetFolder(CatalogPath).DateCreated)
    CurrentYear = DatePart("yyyy", Now)

    Response.Write "<div id=" & Quote & "copyright" & Quote & ">"
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
      Response.Write ", " & LanguageEMail & ": " & "<a href=" & Quote & "mailto:" & CatalogAuthorEmail & Quote & ">" & CatalogAuthorEmail & "</a>"
    end if

    if Len(CatalogAuthorPhone) > 0 then
      Response.Write ", " & LanguageTelephone & ": " & CatalogAuthorPhone
    end if

    Response.Write "</div>"

    ' End of page
    Response.Write "</center>"
    Response.Write "</body>"
    Response.Write "</html>"
  end if
end if
%>
