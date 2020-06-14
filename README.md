# Wallpaper-Changer
Originally created to manually change the desktop wallpaper for Windows XP, this script was later modified to work with Windows 7+. An HTA application was used to provide a user interface to update settings for the script, controlling the source and destination folders.

The VBScript was later modified to work with Windows 7+, which allowed users to have a folder to display rotating wallpapers. The script's main function now was to copy a user-defined number of images from the source folder to the destination folder, allowing the wallpaper slideshow folder to contain seasonally or date-appropriate wallpapers. Windows 7+ didn't support HTA applications any more for security reasons, so there was no way to adjust the user settings except to manually update the "WallpaperChanger Settings.txt" file.

The VB script is being replaced with a Windows Service, which will allow you to keep a master folder of all of your seasonal desktop background images in a single folder. The service will index the files in the master folder and provide an interface that can be used to specify which images should be displayed on a specific date. You can configure how often the service searches for date-appropriate images, and where to draw images from if no date-appropriate images are found. Each time the service searches for a date-appropriate image, it will copy a specified number of images from the source folder to the destination folder. Optionally, you can set up a year-round folder containing images that can be displayed at any time, and set a percentage likelihood that the script will select an image from this folder instead of a date-appropriate image. The user will have to update the settings in Windows so that the desktop wallpaper slideshow points to the destination folder.

The service does not care what the file name is, and will work with JPG, GIF, and PNG images.


