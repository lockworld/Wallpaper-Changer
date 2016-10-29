 ' VBScript File

 Option Explicit

'--------------------------------------------------------------
'                   REFERENCES
'--------------------------------------------------------------
'
' Expected configuration settings file contents should be:
'       configcontents(0) = "Wallpaper Directory:"
'       configcontents(1) = {The configured Wallpaper Directory}
'       configcontents(2) = vbNewLine
'       configcontents(3) = "Current Wallpaper:"
'       configcontents(4) = {The currently selected wallpaper filename}
'       configcontents(5) = vbNewLine
'       configcontents(6) = "Wallpaper Position:"
'       configcontents(7) = {The configured wallpaper position}
'           0 = Center
'           1 = Tile
'           2 = Stretch
'       configcontents(8) = vbNewLine
'       configcontents(9) = "Include 'My Pictures Slideshow?'"
'       configcontents(10) = {Yes/No}
'       configcontents(11) = vbNewLine
'       configcontents(12) = "Wallpaper Last Changed:"
'       configcontents(13) = {Timestamp of last change}
'
'---------------------------------------------------------------
'            END REFERENCES
'---------------------------------------------------------------







 '---------------------------------------------------
 ' Define variables used in script
 '---------------------------------------------------
    Dim _
    colFolders, _
    colSubfolders, _
    configcontents(), _
    configexists, _
    configfilepath, _
    configposition, _
    configslideshow, _
    da, _
    defFile, _
    edate, _
    expLines, _
    extName, _
    file, _
    folderPath, _
    ForAppending, _
    ForReading, _
    ForWriting, _
    foundlines, _
    FSO, _
    i, _
    logcontents, _
    logDirectory, _
    logexists, _
    logFile, _
    logText, _
    max, _
    min, _
    mo, _
    moday, _
    MyFiles, _
    MyFolder, _
    objFile, _
    objFolder, _
    objFSO, _
    objLogFile, _
    objNet, _
    objReadFile, _
    objShell, _
    objStream, _
    objSubfolder, _
    objWallFile, _
    objWMIService, _
    ofolder, _
    oSHApp, _
    scriptPath, _
    sdate, _
    selectedwallpaper, _
    SlideShow, _
    SlideFolder, _
    SPath, _
    strComputer, _
    strDesktop, _
    subFolder, _
    sUserName, _
    sWallPaper, _
    sWinDir, _
    SysFolder, _
    temp, _
    therand, _
    userreply, _
    varPathCurrent, _
    wallDirectory, _
    wallFile, _
    wallText


'-------------------------
' Set script-level variables
'-------------------------
    ' Create the File System Objects
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objNet = CreateObject("WScript.Network")
    Set objShell = CreateObject("WScript.Shell")
    Set oSHApp = CreateObject("Shell.Application")

    strComputer = "."
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


    

    ' Assign the file open/read variables (That won't be changed later in the program)
    ForAppending = 8
    ForReading = 1
    ForWriting = 2 'ForWriting will delete the existing contents before writing to the file

    ' Set the path to the default wallpaper.
    SPath = "C:\Documents and Settings\" & objNet.UserName & "\Local Settings\Application Data\Microsoft"
    
    ' Find the path the the current WallpaperChanger script.
    scriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, WScript.ScriptName) -1)
    

    ' This is the path where the configuration settings file will be found (or created).
    wallDirectory = scriptPath
   

    ' Assigns the filename and path to search for or create the configuration settings file
    wallFile = "WallpaperChanger Settings.txt"
    

    ' Date variables
    mo = Month(Now())
    da = Day(Now())
    if (mo<10) then
        mo = "0" & mo
    else
        mo = "" & mo
    end if
    if (da<10) then
        da = "0" & da
    else
        da = "" & da
    end if

   

On Error Resume Next
Err.Clear

'-------------------------
' Read the contents of the configuration settings file (or create a new file)
'-------------------------
    ' Call a function to read the contents of the configuration settings file (or create one if it does not exist
    '   or if it is invalid.
    GetSetConfigFile()




'-------------------------
' Call the Verify function to determine whether the configuration settings file was built correctly.
'-------------------------
    VerifyConfigSettingsFileLines
        '-------------------------
        ' DEBUGGING: Is there a verified configuration file?
        '    WScript.Echo "Configuration file exists? " & configexists
        '-------------------------




'-------------------------
' If file verifies, call a function to read it. Otherwise, call a function to re-create the file.
'-------------------------
    if (explines = foundlines) then
        ' If the file passes verification, read the file
        ReadWallFile

    else
        ' If size indicates an error in wallFile, call CreateConfigurationSettingsFile function to recreate it
        WScript.Echo "An error has occured with the configuration file: " & vbNewLine & wallDirectory & wallFile _
        & vbNewLine & "Executing built-in pause for 5 seconds to rebuild file."
        WScript.Sleep(2500)

        ' Only if creating it works
        'CreateConfigurationSettingsFile

        ' Otherwise, just modify it
        ModifyConfigurationSettingsfile


    end if

        '-------------------------
        ' DEBUGGING: Display configcontents
        '    i=0
        '    While i<UBound(configcontents)
        '        WScript.Echo configcontents(i)
        '        i=i+1
        '    Wend
        '-------------------------

'-------------------------
' Choose the new wallpaper
'-------------------------
    ' Call a function to randomly select a new wallpaper for the user.
    SelectNewWallpaper


'-------------------------
' Set the new wallpaper
'-------------------------
    ' Call a function to set the user's new wallpaper in the Registry.
    '   NOTE: The new wallpaper will not be displayed until the screen is refreshed (every 3-4 hours, logging in, logging out,
    '         locking the screen, etc.)
    SetUserWallpaper


' Exit the application
WScript.Quit




'------------------------------------------
'     FUNCTIONS
'------------------------------------------

Function GetSetConfigFile ()
    ' This function reads the contents of the Wallpaper Changer Configuration settings file.
    '   If no settings file exists, the user is prompted to create a new configuration settings file.

    '-------------------------
    ' DEBUGGING:
    '    WScript.Echo "Function GetSetConfigFile called"
    '-------------------------

    On Error Resume Next
    Err.Clear

    '---------------------
    ' Look for the configuration settings file.
    '   If no configuration settings file exists, create and configure one now.
    '---------------------
        if objFSO.FolderExists(wallDirectory) then
            Set objFolder = objFSO.GetFolder(wallDirectory)
            ' Folder exists!
            ' This is redundant, as the configuration settings file path is automatically the path to this
            '   script itself. But I left it in in case I decide to allow the user to store the configuration
            '   settings file in a different location someday.
        else
            Set objFolder = objFSO.CreateFolder(wallDirectory)
            ' If the configuration settings file's parent folder does not exist, it will be created here.

            '-------------------------
            ' OPTION: Uncomment the line below if you want the script to notify the user that the folder was created.
            '-------------------------
                WScript.Echo "Successfully installed configuration directory: " & wallDirectory
        
        end if


        ' In either case above, the folder exists. Now, we look for the file itself.
        ' If it doesn't exist, we'll create it now and call the configuration function to prompt the user for input.
        if NOT(objFSO.FileExists(wallDirectory & wallFile)) then
            CreateConfigurationSettingsFile
            ' Call the function that creates the configuration settings file.
            '   NOTE: The CreateConfigurationSettingsFile will, in turn, call a separate script which will allow the user to
            '      review and modify the configuration settings for the program.
            '      Upon successful completion of the script, configexists variable is set to true.

        end if

        
    
    Set objFile = Nothing
    Set objFolder = Nothing

    
End Function







Function VerifyConfigSettingsFileLines ()
    ' This function will read the number of lines in the configuration settings file. If the number of lines in the file
    '   matches the number of expected lines, we can assume the configuration settings file was written correctly.


    '-------------------------
    ' DEBUGGING:
    '    WScript.Echo "Function VerifyConfigSettingsFileLines called"
    '-------------------------

    On Error Resume Next
    Err.Clear
    
    '-------------------------
    ' Set the number of lines you expect to see in the configuration settings file
    '-------------------------
        foundlines = 0
        explines = 14
        '-------------------------
        ' Count the number of lines (i) in the existing configuration settings file
        '-------------------------
            Set objWallFile = objFSO.OpenTextFile (wallDirectory & wallFile, ForReading)
            i = 0
            Do Until objWallFile.AtEndOfStream
            temp = objWallFile.ReadLine
            i=i+1
            Loop
            objWallFile.close
            set objWallFile = Nothing
            
            ' Store the number of lines found into a variable for later use
            foundlines = i
          
            '-------------------------
            ' DEBUGGING:
            '    WScript.Echo "Expected Lines: " & explines & chr(13) & "Found Lines: " & foundlines
            '-------------------------

    

End Function







Function ReadWallFile ()
    ' This function will read the configuration settings file and store each line in the array configcontents.

    '-------------------------
    ' DEBUGGING:
    '    WScript.Echo "Function ReadWallFile called"
    '-------------------------

    On Error Resume Next
    Err.Clear


    ' If configuration settings file exists and passes verification, read it
    Set objWallFile = objFSO.OpenTextFile (wallDirectory & wallFile, ForReading)      
    i=0

    'Save each line into an array variable
    Do Until objWallFile.AtEndOfStream
        Redim Preserve configcontents(i)
        configcontents(i) = objWallFile.ReadLine
        ' msgbox(configcontents(i))
        i=i+1
    Loop



    ' Close the file
    objWallFile.close
    Set objWallFile = Nothing

    if (configcontents(1) > "" AND configcontents(7) > "" AND configcontents(10) > "") then
        configfilepath = configcontents(1)
        configposition = configcontents(7)
        configslideshow = configcontents(10)
        ' If file exists, passes verification, and contains acceptable entries, let the program know it exists
        configexists = true
    else
        configexists = false
    end if

    

End Function






Function CreateConfigurationSettingsFile()
    ' This function will create a new configuration settings file and call a separate script to allow the user to review and modify
    '   the configuration settings for the program.

    On Error Resume Next
    Err.Clear

    Set objWallFile = objFSO.CreateTextFile (wallDirectory & wallFile, ForWriting)
    objWallFile.WriteLine("Wallpaper Directory:")
    objWallFile.WriteLine("C:\Documents and Settings\" & objNet.UserName & "\My Documents\My Pictures")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Current Wallpaper:")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Wallpaper Position:")
    objWallFile.WriteLine("2")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Include My Pictures Slideshow?")
    objWallFile.WriteLine("No")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Wallpaper Last Changed:")
    objWallFile.WriteLine("")
    
    '-------------------------
    ' DEBUGGING:
        WScript.Echo "New configuration file successfully created." & chr(13) & "Function CreateConfigurationSettingsFile called"
    '-------------------------
    


    ModifyConfigurationSettingsFile
    ' Calls a separate script to allow the user to review and modify the configuration settings for the program.
    temp = """" & scriptPath & "\WallpaperChanger_Config.vbs"""
    objShell.Run(temp)
    

End Function







Function ModifyConfigurationSettingsFile()
    ' This function will call an external script to allow the user to review and modify the configuration settings file.

    On Error Resume Next
    Err.Clear




    

End Function







Function SelectNewWallpaper()
    ' This function will find an appropriate wallpaper based on the user's preferred directory and any date-specific 
    ' wallpaper preferences.

    '-------------------------
    ' DEBUGGING:
    '    WScript.Echo "Function SelectNewWallpaper called"
    '-------------------------

    On Error Resume Next
    Err.Clear
   
   

    if (configexists = true) then
        ' Verify that the selected wallpaper directory actually exists.
        '   If not, then give user the option to adjust settings or quit program
        If objFSO.FolderExists(configfilepath) then
            Set SysFolder = FSO.GetFolder(configfilepath)
        else
            userreply = msgbox("Unable to find the selected wallpaper directory. Would you like to change" _
            & " your wallpaper changer settings now?", vbYesNo)

            if (userreply = 6) then
                ' If user agrees, call the configuration script to modify settings.
                ModifyConfigurationSettingsFile
                
            end if
            
            ' Quit this script (will be re-run if/when user modifies configuration settings file).
            WScript.Quit
            
        end if

        '-------------------------
        ' DEBUGGING: Notify user that wallpaper directory is valid
        '    WScript.Echo "wallDirectory is valid!" & chr(13) & configfilepath
        '-------------------------

        '-------------------------
        ' Set the default folder for images (in case there are no override folders or override folders are empty
        '-------------------------
        Set MyFolder = FSO.GetFolder(configfilepath)
        folderPath = configfilepath



        '-------------------------
        ' Set a ((1-({min/max}))*100)% chance that the program will check for special "override" folder for a particular month or a particular date range 
        ' If min=1 and max=4, you have a 75% probability that the program will check for the "override" folder instead of the default folder
        ' If max=1000, you have a nearly 100% probability that the program will ONLY look in the "override" folder and NEVER in the default folder
        ' 
        '-------------------------
            max = 1000
            min = 1
            Randomize
            therand = Int((max-min+1) * Rnd+min)

            if (therand<max) then
                ' If a folder exists for a specific month...
                    if (objFSO.FolderExists(configfilepath & "\" & mo)) then
                        ' If no images in this folder, use the default folder
                        Set temp = FSO.GetFolder(configfilepath & "\" & mo)
                        if (temp.Files.Count > 0) then
                            Set MyFolder = FSO.GetFolder(configfilepath & "\" & mo)
                            folderPath = configfilepath & "\" & mo
                            moday = mo
                        end if
                    end if
        
            

                ' If a folder exists for a specific date range...
                    Set objFolder = objFSO.GetFolder(configfilepath)
                    Set colSubfolders = objFolder.Subfolders
                    For Each objSubfolder in colSubfolders
                        if (instr(1,objSubfolder.Name,"-") > 0) then
                            sdate = left(objSubfolder.Name, instr(1, objSubfolder.Name,"-")-1)
                            edate = right(objSubfolder.Name, InstrRev(objSubfolder.Name, "-")-1)
                            if ((mo & "_" & da) >= sdate AND (mo & "_" & da) <= edate) then
                                ' If no images in this folder, use the default folder
                                Set temp = FSO.GetFolder(configcontents(1) & "\" & objSubfolder.Name)
                                if (temp.Files.Count > 0) then
                                    Set MyFolder = FSO.GetFolder(configcontents(1) & "\" & objSubfolder.Name)
                                    folderPath = configcontents(1) & "\" & objSubfolder.Name
                                    moday = objSubfolder.Name 
                                end if
                            else
                                '-------------------------
                                ' DEBUGGING: Show the dates calculated
                                '   WScript.Echo "mo_da = " & mo & "_" & da & chr(13) & "sdate = " & sdate
                                '-------------------------
                            end if
                        end if
                    Next
            end if

            
        '-------------------------
        ' If a folder exists and contains images for a specific day, ALWAYS use that folder 
        '-------------------------
            if (objFSO.FolderExists(configfilepath & "\" & mo & "_" & da)) then
                ' If no images in this folder, use the default folder
                Set temp = FSO.GetFolder(configfilepath & "\" & mo & "_" & da)
                if (temp.Files.Count > 0) then
                    Set MyFolder = FSO.GetFolder(configfilepath & "\" & mo & "_" & da)
                    folderPath = configcontents(1) & "\" & mo & "_" & da
                    moday = mo & "_" & da
                end if
            end if

        
        '-------------------------
        ' DEBUGGING: Tell user folder exists for this month
        '        WScript.Echo "You are using the following folder: " & folderPath
        '-------------------------

        '-------------------------
        ' Select a random picture to use for the wallpaper
        '-------------------------
            max = MyFolder.Files.Count
            min = 1
            Randomize
            therand = Int((max-min+1) * Rnd+min)

            temp = ""
            i = 0


            ' Select a file with qualifying extension
            For each file in MyFolder.Files
                i = i+1
                extName = right(file.Name, 3)
                if (extName="jpg" OR extName="JPG" OR extName="bmp" OR extName="BMP" OR extName="gif" OR extName="GIF" OR extName="png" OR extName="PNG") then
                    ' Get first file as default file
                    if (temp="") then
                        temp = file.Name
                        defFile = file.Name
                    end if

                    if (i=therand) then
                        temp = file.Name
                    end if
                else
                    if (i=therand) then
                        min=i+1
                        Randomize
                        therand = Int((max-min+1)*Rnd+min)
                    end if
                end if

            Next

            if (temp="") then
                selectedwallpaper = defFile
            else
                selectedwallpaper = temp
            end if


    else
        '-------------------------
        ' If no configuration file has been found, but this subroutine has been called...exit application
        '-------------------------
        WScript.Quit
        ' Application has failed.
    end if
    

'-------------------------END OF SELECT NEW WALLPAPER FUNCTION-------------------------
End Function





Function SetUserWallpaper()
    ' This function changes the registry settings to select the new wallpaper and other user preferences from the 
    '   configuration settings file.

    On Error Resume Next
    Err.Clear

    if (isnull(selectedwallpaper) OR selectedwallpaper = "") then
        msgbox("No wallpaper found in " & folderPath & "\" & selectedwallpaper)
    else
        ' Remove existing wallpaper file(s)
        if objFSO.FileExists(SysFolder & "\Wallpaper1.bmp") then
            objFSO.DeleteFile SysFolder & "\Wallpaper1.bmp"
        end if

        if objFSO.FileExists(SysFolder & "\Wallpaper1.jpg") then
            objFSO.DeleteFile SysFolder & "\Wallpaper1.jpg"
        end if

        if objFSO.FileExists(SysFolder & "\Wallpaper1.gif") then
            objFSO.DeleteFile SysFolder & "\Wallpaper1.gif"
        end if

        if objFSO.FileExists(SysFolder & "\Wallpaper1.png") then
            objFSO.DeleteFile SysFolder & "\Wallpaper1.png"
        end if

        objFSO.CopyFile folderPath & "\" & selectedwallpaper , SPath & "\" & "Wallpaper1." & right(selectedwallpaper, 3)

        ' Update the settings file
        Set objWallFile = objFSO.CreateTextFile (wallDirectory & wallFile, ForWriting)

        objWallFile.WriteLine("Wallpaper Directory:")
        objWallFile.WriteLine(configfilepath)
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Current Wallpaper:")
        if (configfilepath = folderPath) then
            objWallFile.WriteLine(selectedwallpaper)
        else
            objWallFile.WriteLine(moday & "\" & selectedwallpaper)
        end if
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Wallpaper Position:")
        objWallFile.WriteLine(configposition)
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Include 'My Pictures Slideshow?'")
        objWallFile.WriteLine(configslideshow)
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Wallpaper Last Changed:")
        objWallFile.WriteLine(Now())

        objWallFile.close
        Set objWallFile = Nothing

        ' Set the selected wallpaper as the Windows desktop wallpaper
        sWallPaper = SPath & "\" & "Wallpaper1." & right(temp, 3)
  
        ' update in registry
        objShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper
        if (configposition=1) then
            objShell.RegWrite "HKCU\Control Panel\Desktop\TileWallpaper", 1
        else
            objShell.RegWrite "HKCU\Control Panel\Desktop\TileWallpaper", 0
        end if
        if (wallPosition > -1 AND wallPosition < 3) then
           objShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", configposition
        else
           objShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", 2
        end if
            
        if (configslideshow="Yes") then
            objShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Screen Saver.Slideshow\ImageDirectory", FolderPath
            '-------------------------
            ' DEBUGGING: Notify user of My Pictures Slideshow folder
            '   MsgBox("Your My Pictures Slideshow is set to: " & FolderPath)
            '-------------------------
        end if
        ' let the system know about the change
        objShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, True
      

    end if

End Function






