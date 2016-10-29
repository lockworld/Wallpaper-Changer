# Wallpaper-Changer
Originally created to manually change the desktop wallpaper for Windows XP, this script was later modified to work with Windows 7+.

This is the newer version of the script meant to work with Windows 7+. It is incomplete...there is no way to adjust the user settings except to manually update the "WallpaperChanger Settings.txt" file.

This script allows you to keep a master folder of all of your seasonal desktop background images in a single folder. You can create subfolders named with a date-range as shown in the example below. When the script runs, it looks for a folder with a date range containing the current date. If found, the script will use this folder as the source folder for the images and copy a new image from the source folder to a destination folder. The script also allows a 20% chance that an image from the main wallpaper folder will be selected instead of an image from the seasonal folder, allowing some images to be displayed year-round, while most images are displayed only as seasonably appropriate. The user can then set up his desktop wallpaper slideshow to point to the destination folder.

Users will set up a scheduled task to run this script on a regular basis. Each time the script is run, it will select one image from the source folder and copy it to the destination folder. The script will only keep as many images in the destination folder as configured in the settings file.

Once everything is set up, the script just runs in the background, and the user's selection of desktop wallpapers changes throughout the year.

For example, to display winter-themed desktop backgrounds from January through March, name your folder:

01_01-03_31

This is parsed into the date range of January 1 through March 31.



To display Christmas backgrounds between December 1 and December 25, name your folder:

12_01-12_25

This is parsed into the date range of December 1 through December 25.



To display select wallpapers only on a specific date (such as an anniversary or memorial), name your folder as follows:

09_11

The contents of this folder will only be displayed on September 11.



If you want certain images to be available on a limited basis year-round, just put those images into the main wallpaper folder.

The script does not care what the file name is, and will work with JPG, GIF, and PNG images.

I recommend setting the scheduled task to run at least as many times a day as you have set for maximum images because only one new image is processed each time the task is run, and only one of the older iamges will be removed each time the task is run.


