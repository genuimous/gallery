# Gallery
ASP web-based multi-user photo gallery

If wiaaut.dll not exists in %SystemRoot%\system32, need to install it from wiaautsdk.zip (see ReadMe.txt inside the archive)

To install, just copy this files to web directory on your server:
- album.asp
- catalog.asp
- image.asp
- album.cfg
- catalog.cfg
- image.cfg
- imaging.inc
- utils.inc
Don't forget give ASP execition for this dir!

After that, create (virtual) subdir with name "catalog" - pictures will be stored here. Now, it possible to create dir "my_photos" in web directory "<gallery_app_dir>\catalog" (full path will be just like that: "<gallery_app_dir>\catalog\my_photos") and place here subdirs with ypor photos. Only JPEG (.jpg, .jpeg) pictures supported. For example:
 - <gallery_app_dir>\catalog\my_photos\album_1
 - <gallery_app_dir>\catalog\my_photos\album_2
 - ...
 - <gallery_app_dir>\catalog\my_photos\album_n

Note: you need allow write access for ASP process to <gallery_app_dir>\catalog and subdirs because server will generate thumbnails.

Now, enter in web browser:
> http://your_host/gallery_app_dir/catalog.asp?name=my_photos 

Gallery web page with list of your albums should appear. You can create as many dirs as you want (i.e. "my_photos", "your_photos", "theyr_photos" etc.) and query them using name parameter in the URL.
  
Also, there are catalog.cfg and album.cfg files for custom data:
- catalog.cfg file in <gallery_app_dir> provides default labels for new catalogs, but you can copy in to your catalog subdir and redefine there (catalog.cfg found in "<gallery_app_dir>/catalog/<subdir>" has more priority then found in "<gallery_app_dir>")
- album.cfg file in <gallery_app_dir> provides default labels for new albums, but you can copy in to your album subdir and redefine there (album.cfg found in "<gallery_app_dir>/catalog/<subdir>/<album_subdir>" has more priority then found in "<gallery_app_dir>")
  
For example, complete structure will be like that:
 - <gallery_app_dir>\catalog\my_photos\album_1 - album 1 in my_photos
 - <gallery_app_dir>\catalog\my_photos\album_1\album.cfg - optional config for album 1 in my_photos
 - <gallery_app_dir>\catalog\my_photos\album_2 - album 2 in my_photos
 - ...
 - <gallery_app_dir>\catalog\my_photos\album_n - album N in my_photos
 - <gallery_app_dir>\catalog\my_photos\catalog.cfg - optional config for my_photos
 - ...
 - <gallery_app_dir>\catalog\your_photos\album_1 - album 1 in your_photos
 - <gallery_app_dir>\catalog\your_photos\album_1\album.cfg - optional config for album 1 in your_photos
 - <gallery_app_dir>\catalog\your_photos\album_2 - album 2 in your_photos
 - ...
  
Parameters for catalog.cfg:
- language (from <gallery_app_dir>\language)
- style (from <gallery_app_dir>\style)
- title
- logo_size
- author
- author_site
- author_email
- author_phone

Parameters for album.cfg:
- logo (the "logo" filename used for album in catalog, default is "logo.jpg")
- title
- description
- thumbnail_size
- thumbnail_grid_cols
- thumbnail_grid_rows
- thumbnail_indent

See original catalog.cfg and album.cfg in <gallery_app_dir> for reference.
