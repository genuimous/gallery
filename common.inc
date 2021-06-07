<%
Quote = Chr(34)
OSDirDelimiter = "\"
WebDirDelimiter = "/"
Copyright = "&copy;"

ImageExtensions = array(".jpg", ".jpeg")

RootPath = Server.MapPath(".")

DefaultCatalogDefense = "no"
DefaultCatalogLanguage = ""
DefaultCatalogStyle = "grey"
DefaultCatalogLogoSize = 128

DefaultAlbumLogoFileName = "logo.jpg"
DefaultAlbumArchiveFileName = "album.zip"
DefaultAlbumThumbnailSize = 240
DefaultAlbumThumbnailGridCols = 3
DefaultAlbumThumbnailGridRows = 3
DefaultAlbumThumbnailIndent = 10

DefaultImagePreviewSize = 960
DefaultImageNavigation = True

CatalogDir = "catalog"
CatalogConfigFileName = "catalog.cfg"
CatalogInfoFileName = "info.txt"
CatalogDefenseKey = "defense"
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
AlbumArchiveFileNameKey = "archive"
AlbumTitleKey = "title"
AlbumDescriptionKey = "description"
AlbumThumbnailSizeKey = "thumbnail_size"
AlbumThumbnailGridColsKey = "thumbnail_grid_cols"
AlbumThumbnailGridRowsKey = "thumbnail_grid_rows"
AlbumThumbnailIndentKey = "thumbnail_indent"

ImageConfigFileName = "image.cfg"
ImagePreviewSizeKey = "preview_size"
ImageNavigationKey = "navigation"

ThumbnailPostfix = ".thumbnail"
PreviewPostfix = ".preview"

LanguageCharset = "iso-8859-1"
LanguageThumbnailHint = "Click image for enlarge"
LanguagePreviewHint = "Click image for download"
LanguageMainPage = "Main page"
LanguageMainPageTitle = "To main page"
LanguageDownloadAll = "Download all"
LanguageDownloadAllTitle = "Download the whole album"
LanguageIndex = "Index"
LanguageIndexTitle = "Go to index"
LanguageIndexNumTitle = "Jump to index page #"
LanguagePrevious = "Previous"
LanguagePreviousTitle = "Previous photo"
LanguageNext = "Next"
LanguageNextTitle = "Next photo"
LanguageEMail = "e-mail"
LanguageTelephone = "tel."
LanguagePrivateText = "Provide album name for access"
LanguagePrivateAction = "Proceed"

StyleDir = "style"

set FileSystem = CreateObject("Scripting.FileSystemObject")

sub LoadLanguage(byval Language)
	if Len(Language) > 0 then
		LanguageConfig = GetFileContent(Server.MapPath("language") & OSDirDelimiter & Language & ".cfg")

		LanguageCharset = GetConfigValue(LanguageConfig, "charset", LanguageCharset)
		LanguageThumbnailHint = GetConfigValue(LanguageConfig, "thumbnail_hint", LanguageThumbnailHint)
		LanguagePreviewHint = GetConfigValue(LanguageConfig, "preview_hint", LanguagePreviewHint)
		LanguageMainPage = GetConfigValue(LanguageConfig, "main_page", LanguageMainPage)
		LanguageMainPageTitle = GetConfigValue(LanguageConfig, "main_page_title", LanguageMainPageTitle)
		LanguageDownloadAll = GetConfigValue(LanguageConfig, "download_all", LanguageDownloadAll)
		LanguageDownloadAllTitle = GetConfigValue(LanguageConfig, "download_all_title", LanguageDownloadAllTitle)
		LanguageIndex = GetConfigValue(LanguageConfig, "index", LanguageIndex)
		LanguageIndexTitle = GetConfigValue(LanguageConfig, "index_title", LanguageIndexTitle)
		LanguageIndexNumTitle = GetConfigValue(LanguageConfig, "index_num_title", LanguageIndexNumTitle)
		LanguagePrevious = GetConfigValue(LanguageConfig, "previous", LanguagePrevious)
		LanguagePreviousTitle = GetConfigValue(LanguageConfig, "previous_title", LanguagePreviousTitle)
		LanguageNext = GetConfigValue(LanguageConfig, "next", LanguageNext)
		LanguageNextTitle = GetConfigValue(LanguageConfig, "next_title", LanguageNextTitle)
		LanguageEMail = GetConfigValue(LanguageConfig, "e_mail", LanguageEMail)
		LanguageTelephone = GetConfigValue(LanguageConfig, "telephone", LanguageTelephone)
		LanguagePrivateText = GetConfigValue(LanguageConfig, "private_text", LanguagePrivateText)
		LanguagePrivateAction = GetConfigValue(LanguageConfig, "private_action", LanguagePrivateAction)
	end if
end sub

function GetConfigValue(byval Config, byval Key, byval DefaultValue)
	GetConfigValue = DefaultValue

	dim ConfigLines
	ConfigLines = Split(Config, vbCrLf)

	for each ConfigLine in ConfigLines
		dim Associations
		Associations = Split(ConfigLine, "=")

		if UBound(Associations) = 1 then
			dim ConfigKey
			dim ConfigValue

			ConfigKey = Trim(Associations(0))
			ConfigValue = Replace(Trim(Associations(1)), Quote, "")

			if ConfigKey = Key then
				GetConfigValue = ConfigValue
				exit for
			end if
		end if
	next
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

	set FileSystem = nothing
end function

function ContainsValue(byval ValueList, byval Choice)
	ContainsValue = False

	for each Value in ValueList
		if Lcase(Value) = Lcase(Choice) then
			ContainsValue = True
			exit for
		end if
	next
end function

function ClearText(byval Text)
	ClearText = Trim(Replace(Text, vbCrLf, ""))

	do while InStr(1, ClearText, "	")
		CleanText = Replace(ClearText, "	", " ")
	loop
end function
%>
