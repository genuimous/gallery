<%
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
			ConfigValue = Replace(Trim(Associations(1)), Chr(34), "")

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
