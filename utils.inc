<%
Quote = Chr(34)
OSDirDelimiter = "\"
WebDirDelimiter = "/"
Copyright = "&copy;"

ImageExtensions = array(".jpg", ".jpeg")

RootPath = Server.MapPath(".")

DefaultCatalogLanguage = ""
DefaultCatalogStyle = "grey"
DefaultCatalogLogoSize = 128

DefaultAlbumLogoFileName = "logo.jpg"
DefaultAlbumThumbnailSize = 240
DefaultAlbumThumbnailGridCols = 3
DefaultAlbumThumbnailGridRows = 3
DefaultAlbumThumbnailIndent = 10

DefaultImageNavigation = True

CatalogDir = "catalog"
CatalogConfigFileName = "catalog.cfg"
CatalogIntroTextFileName = "intro.txt"
CatalogOutroTextFileName = "outro.txt"
CatalogLanguageKey = "language"
CatalogStyleKey = "style"
CatalogTitleKey = "title"
CatalogLogoSizeKey = "logo_size"
CatalogAuthorKey = "author"
CatalogAuthorSiteKey = "author_site"
CatalogAuthorEmailKey = "author_email"
CatalogAuthorPhoneKey = "author_phone"

AlbumConfigFileName = "album.cfg"
AlbumLogoFileNameKey = "logo"
AlbumTitleKey = "title"
AlbumDescriptionKey = "description"
AlbumThumbnailSizeKey = "thumbnail_size"
AlbumThumbnailGridColsKey = "thumbnail_grid_cols"
AlbumThumbnailGridRowsKey = "thumbnail_grid_rows"
AlbumThumbnailIndentKey = "thumbnail_indent"

ImageConfigFileName = "image.cfg"
ImageNavigationKey = "navigation"

ThumbnailPostfix = ".thumbnail"

LanguageCharset = "iso-8859-1"
LanguageThumbnailHint = "Click image for enlarge"
LanguageMainPage = "Main page"
LanguageMainPageTitle = "To main page"
LanguageIndex = "Index"
LanguageIndexTitle = "Back to index"
LanguageIndexNumTitle = "Jump to index page #"
LanguagePrevious = "Previous"
LanguagePpreviousTitle = "Previous photo"
LanguageNext = "Next"
LanguageNextTitle = "Next photo"
LanguageEMail = "e-mail"
LanguageTelephone = "tel."

StyleDir = "style"

set FileSystem = CreateObject("Scripting.FileSystemObject")

sub LoadLanguage(byval Language)
  if Len(Language) > 0 then
    LanguageConfig = GetFileContent(Server.MapPath("language") & OSDirDelimiter & Language & ".cfg")

    LanguageCharset = GetConfigValue(LanguageConfig, "charset", LanguageCharset)
    LanguageThumbnailHint = GetConfigValue(LanguageConfig, "thumbnail_hint", LanguageThumbnailHint)
    LanguageMainPage = GetConfigValue(LanguageConfig, "main_page", LanguageMainPage)
    LanguageMainPageTitle = GetConfigValue(LanguageConfig, "main_page_title", LanguageMainPageTitle)
    LanguageIndex = GetConfigValue(LanguageConfig, "index", LanguageIndex)
    LanguageIndexTitle = GetConfigValue(LanguageConfig, "index_title", LanguageIndexTitle)
    LanguageIndexNumTitle = GetConfigValue(LanguageConfig, "index_num_title", LanguageIndexNumTitle)
    LanguagePrevious = GetConfigValue(LanguageConfig, "previous", LanguagePrevious)
    LanguagePpreviousTitle = GetConfigValue(LanguageConfig, "previous_title", LanguagePpreviousTitle)
    LanguageNext = GetConfigValue(LanguageConfig, "next", LanguageNext)
    LanguageNextTitle = GetConfigValue(LanguageConfig, "next_title", LanguageNextTitle)
    LanguageEMail = GetConfigValue(LanguageConfig, "e_mail", LanguageEMail)
    LanguageTelephone = GetConfigValue(LanguageConfig, "telephone", LanguageTelephone)
  end if
end sub

function RoundUp(byval x)
  if x - Round(x) > 0 then
    RoundUp = Round(x) + 1
  else
    RoundUp = Round(x)
  end if
end function

function GetConfigValue(byval Config, byval Key, byval DefaultValue)
  GetConfigValue = SeparateField(Config, Key & "=" & Chr(34), Chr(34) & vbCrLf)
  if Len(GetConfigValue) = 0 then GetConfigValue = DefaultValue
end function

function SeparateField(byval SourceString, byval StartString, byval EndString)
  dim StartPos, EndPos

  StartPos = InStr(1, SourceString, StartString, vbTextCompare)

  if StartPos > 0 Then
    StartPos = StartPos + Len(StartString)
    EndPos = InStr(StartPos, SourceString, EndString, vbTextCompare)

    if EndPos > 0 then
      SeparateField = Trim(Mid(SourceString, StartPos, EndPos - StartPos))
    end if
  end if
end function

function GetFileContent(byval FileName)
  dim FileSystem

  set FileSystem = CreateObject("Scripting.FileSystemObject")

  if FileSystem.FileExists(FileName) then
    if FileSystem.GetFile(FileName).Size > 0 then
      GetFileContent = FileSystem.OpenTextFile(FileName).ReadAll
    else
      GetFileContent = ""
    end if
  else
    GetFileContent = ""
  end if

  set FileSystem = Nothing
end function

function ContainsValue(byval ValueList, byval Choice)
  ContainsValue = False

  for each Value in ValueList
    if Value = Choice then
      ContainsValue = True
      exit for
    end if
  next
end function
%>
