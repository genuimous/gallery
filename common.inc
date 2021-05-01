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

StyleDir = "style"

set FileSystem = CreateObject("Scripting.FileSystemObject")
%>
