<%
sub GenerateImagePreview(byval ImageFileName, byval PreviewFileName, byval Size, byval Square)
	on error resume next

	dim ImageFile
	dim ImageCropProcess
	dim ImageResizeProcess
	dim LeftCrop, RightCrop, TopCrop, BottomCrop
	dim AspectRatio
	dim PreviewWidth, PreviewHeight

	set ImageFile = CreateObject("WIA.ImageFile")

	ImageFile.LoadFile ImageFileName

	if Square then
		if ImageFile.Width <> ImageFile.Height then
			if ImageFile.Width > ImageFile.Height then
				LeftCrop = Round((ImageFile.Width - ImageFile.Height) / 2)
				RightCrop = ImageFile.Width - ImageFile.Height - LeftCrop
				TopCrop = 0
				BottomCrop = 0
			else
				LeftCrop = 0
				RightCrop = 0
				TopCrop = Round((ImageFile.Height - ImageFile.Width) / 2)
				BottomCrop = ImageFile.Height - ImageFile.Width - TopCrop
			end if

			set ImageCropProcess = CreateObject("WIA.ImageProcess")

			with ImageCropProcess
				.Filters.Add .FilterInfos("Crop").FilterID

				with .Filters.Item(1).Properties
					.Item("Left") = LeftCrop
					.Item("Right") = RightCrop
					.Item("Top") = TopCrop
					.Item("Bottom") = BottomCrop
				end with
			end with

			set ImageFile = ImageCropProcess.Apply(ImageFile)

			set ImageCropProcess = nothing
		end if
	end if

	if ImageFile.Width - Size > 0 or ImageFile.Height - Size > 0 then
		AspectRatio = ImageFile.Width / ImageFile.Height

		if AspectRatio >= 1 then 
			PreviewWidth = Size
			PreviewHeight = Round(Size / AspectRatio)
		else
			PreviewWidth = Round(Size / AspectRatio)
			PreviewHeight = Size
		end if

		set ImageResizeProcess = CreateObject("WIA.ImageProcess")

		with ImageResizeProcess
			.Filters.Add .FilterInfos("Scale").FilterID

			with .Filters.Item(1).Properties
				.Item("MaximumWidth") = PreviewWidth
				.Item("MaximumHeight") = PreviewHeight
			end with
		end with

		set ImageFile = ImageResizeProcess.Apply(ImageFile)

		set ImageResizeProcess = nothing
	end if

	ImageFile.SaveFile PreviewFileName

	set ImageFile = nothing
end sub
%>
