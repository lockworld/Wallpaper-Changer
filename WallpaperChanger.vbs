 ' VBScript File

 Option Explicit

'--------------------------------------------------------------
'                   REFERENCES
'--------------------------------------------------------------
'
' Expected configuration settings file contents should be:
'       configcontents(0) = "Get New Images From This Directory:"
'       configcontents(1) = {The configured Wallpaper Source Directory}
'       configcontents(2) = vbNewLine
'       configcontents(3) = "Store Current Wallpaper Images in This Directory:"
'       configcontents(4) = {The configured Wallpaper Destination Directory}
'       configcontents(5) = vbNewLine
'       configcontents(6) = "Keep This Number of Wallpaper Images"
'       configcontents(7) = {The number of images to store}
'       configcontents(8) = vbNewLine
'       configcontents(9) = "Wallpaper Last Changed:"
'       configcontents(10) = {Timestamp of last change}
'
'
' Log Text Sample
' Message type is bracket + 20 characters + bracket
' [tab]
' Timestamp is bracket + 14 characters + bracket
' [tab]
' Message
'
'       [Message Type 20 Char] [timestamp]      Message
'---------------------------------------------------------------
'            END REFERENCES
'---------------------------------------------------------------







 '---------------------------------------------------
 ' Define variables used in script
 '---------------------------------------------------
' Dim
' colFolders, _


    Dim _
    colSubfolders, _
    configcontents(), _
    configsourcepath, _
    configdestpath, _
    configkeepimages, _
    configkeepimages_orig, _
    configexists, _
    da, _
    defFile, _ 
    DestFolder, _
    DestinationFolder, _
    edate, _
    error, _
    errorcount, _
    expLines, _
    extName, _
    file, _
    folderPath, _
    ForAppending, _
    ForReading, _
    ForWriting, _
    foundlines, _
    FSO, _
    ho, _
    i, _
    j, _
    logcontents, _
    logDirectory, _
    logexists, _
    logFile, _
    logText, _
    max, _
    mi, _
    min, _
    mo, _
    moday, _
    MyFiles, _
    NewFileName, _
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
    scriptDirectory, _
    scriptPath, _
    sdate, _
    se, _
    selectedwallpaper, _
    SlideShow, _
    SlideFolder, _
    SourceFolder, _
    SPath, _
    strComputer, _
    strDesktop, _
    subFolder, _
    sUserName, _
    sWallPaper, _
    sWinDir, _
    temp, _
    temp1, _
    temp2, _
    therand, _
    timestamp, _
    userreply, _
    varPathCurrent, _
    wallFile, _
    wallText, _
    ye


'-------------------------
' Set script-level variables
'-------------------------
    errorcount=0
    
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

    

    ' Find the path the the current WallpaperChanger script.
    scriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, WScript.ScriptName) -1)
    

    ' This is the path where the configuration settings file will be found (or created).
    scriptDirectory = scriptPath
   

    ' Assigns the filename and path to search for or create the configuration settings file
    wallFile = "WallpaperChanger Settings.txt"
    

    ' Date variables
    ye = Year(Now())
    mo = Month(Now())
    da = Day(Now())
    ho = Hour(Now())
    mi = Minute(Now())
    se = Second(Now())
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
    if (ho<10) then
        ho = "0" & ho
    else
        ho = "" & ho
    end if
    if (mi<10) then
        mi = "0" & mi
    else
        mi = "" & mi
    end if
    if (se<10) then
        se = "0" & se
    else
        se = "" & se
    end if
    
    timestamp = ye & mo & da & ho & mi & se

   

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
        WScript.Echo "An error has occured with the configuration file: " & vbNewLine & scriptDirectory & wallFile _
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
    selectedwallpaper = SelectNewWallpaper
    
'-------------------------
' Set the new wallpaper
'-------------------------
    ' Copy the new wallpaper into the destination wallpaper folder
    CopyNewWallpaper



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
    '   If no configuration settings file exists, call another function to create and configure one now.
    '---------------------
        if objFSO.FolderExists(scriptDirectory) then
            Set objFolder = objFSO.GetFolder(scriptDirectory)
            ' Folder exists!
            ' This is redundant, as the configuration settings file path is automatically the path to this
            '   script itself. But I left it in in case I decide to allow the user to store the configuration
            '   settings file in a different location someday.
        else
            Set objFolder = objFSO.CreateFolder(scriptDirectory)
            ' If the configuration settings file's parent folder does not exist, it will be created here.

            '-------------------------
            ' OPTION: Uncomment the line below if you want the script to notify the user that the folder was created.
            '-------------------------
            '    WScript.Echo "Successfully installed configuration directory: " & scriptDirectory
        
        end if


        ' In either case above, the folder exists. Now, we look for the file itself.
        ' If it doesn't exist, we'll create it now and call the configuration function to prompt the user for input.
        if NOT(objFSO.FileExists(scriptDirectory & wallFile)) then
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
        explines = 11
        '-------------------------
        ' Count the number of lines (i) in the existing configuration settings file
        '-------------------------
            Set objWallFile = objFSO.OpenTextFile (scriptDirectory & wallFile, ForReading)
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
    Set objWallFile = objFSO.OpenTextFile (scriptDirectory & wallFile, ForReading)
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

    if (configcontents(1) > "" AND configcontents(4) > "" AND configcontents(7) > "") then
        ' Set the path to the wallpaper source folder
        configsourcepath = configcontents(1)

        ' Set the path to the wallpaper destination folder
        configdestpath = configcontents(4)
        
        ' Set the number of images to keep
        configkeepimages = configcontents(7)
        if (configkeepimages="" OR configkeepimages="0") then
           configkeepimages="10"
        end if
        if (CInt(configkeepimages)<10) then
           configkeepimages = "0" & configkeepimages
        end if
        
        
        configkeepimages_orig = CInt(configkeepimages)
        
        '        configposition = configcontents(7)
        '        configslideshow = configcontents(10)
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

    '-------------------------
    ' DEBUGGING:
    '    WScript.Echo "Function CreateConfigurationSettingsFile called"
    '-------------------------

    Set objWallFile = objFSO.CreateTextFile (scriptDirectory & wallFile, ForWriting)
    objWallFile.WriteLine("Get New Images From This Directory:")
    objWallFile.WriteLine("C:\Users\" & objNet.UserName & "\Pictures\Wallpaper_Source")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Store Current Wallpaper Images in This Directory:")
    objWallFile.WriteLine("C:\Users\" & objNet.UserName & "\Pictures\Wallpaper_Destination")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Keep This Number of Wallpaper Images:")
    objWallFile.WriteLine("5")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Wallpaper Last Changed:")
    objWallFile.WriteLine("")
    
    '-------------------------
    ' DEBUGGING:
    '    WScript.Echo "New configuration file successfully created." & chr(13) & "Function ModifyConfigurationSettingsFile called"
    '-------------------------
    


    ModifyConfigurationSettingsFile
    
    

End Function







Function ModifyConfigurationSettingsFile()
    ' This function will call an external script to allow the user to review and modify the configuration settings file.

    On Error Resume Next
    Err.Clear
    '-------------------------
    ' DEBUGGING:
    '    WScript.Echo "Function ModifyConfigurationSettingsFile called"
    '-------------------------
    ' Calls a separate script to allow the user to review and modify the configuration settings for the program.
    temp = """" & scriptPath & "\WallpaperChanger_Config.vbs"""
    WScript.Echo "This is where the old wallpaper changer config would run...currently disabled"
    'objShell.Run(temp)


    

End Function







Function SelectNewWallpaper()
    ' This function will find an appropriate wallpaper based on the user's preferred directory and any date-specific 
    ' wallpaper preferences and copy it to the destination directory.

    '-------------------------
    ' DEBUGGING:
    '    WScript.Echo "Function SelectNewWallpaper called"
    '-------------------------

    On Error Resume Next
    Err.Clear
   
   

    if (configexists = true) then
        ' Verify that the selected wallpaper directory actually exists.
        '   If not, then give user the option to adjust settings or quit program
        If (objFSO.FolderExists(configsourcepath) AND objFSO.FolderExists(configdestpath)) then
            Set SourceFolder = FSO.GetFolder(configsourcepath)
            Set DestinationFolder = FSO.GetFolder(configdestpath)
        else
            userreply = msgbox("Unable to find the selected wallpaper directory. Would you like to change" _
            & " your wallpaper changer settings now?", vbYesNo)

            if (userreply = 6) then
                ' If user agrees, call the configuration script to modify settings.
                ModifyConfigurationSettingsFile
                ' Quit this script (will be re-run if/when user modifies configuration settings file).
                WScript.Quit
            end if
        end if

        '-------------------------
        ' DEBUGGING: Notify user that wallpaper directory is valid
        '    WScript.Echo "SourceDirectory is valid!" & chr(13) & configsourcepath & chr(13) & "DestinationDirectory is valid!" & chr(13) & configdestpath
        '-------------------------

        '-------------------------
        ' Set the default folder for source images (in case there are no override folders or override folders are empty
        '-------------------------
        Set SourceFolder = FSO.GetFolder(configsourcepath)
        folderPath = configsourcepath



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
                    if (objFSO.FolderExists(configsourcepath & "\" & mo)) then
                        ' If no images in this folder, use the default folder
                        Set temp = FSO.GetFolder(configsourcepath & "\" & mo)
                        if (temp.Files.Count > 0) then
                            Set SourceFolder = FSO.GetFolder(configsourcepath & "\" & mo)
                            folderPath = configsourcepath & "\" & mo
                            moday = mo
                        end if
                    end if
        
            

                ' If a folder exists for a specific date range...
                    Set objFolder = objFSO.GetFolder(configsourcepath)
                    Set colSubfolders = objFolder.Subfolders
                    For Each objSubfolder in colSubfolders
                        if (instr(1,objSubfolder.Name,"-") > 0) then
                            sdate = left(objSubfolder.Name, instr(1, objSubfolder.Name,"-")-1)
                            edate = right(objSubfolder.Name, InstrRev(objSubfolder.Name, "-")-1)
                            if ((mo & "_" & da) >= sdate AND (mo & "_" & da) <= edate) then
                                ' If no images in this folder, use the default folder
                                Set temp = FSO.GetFolder(configcontents(1) & "\" & objSubfolder.Name)
                                if (temp.Files.Count > 0) then
                                    Set SourceFolder = FSO.GetFolder(configcontents(1) & "\" & objSubfolder.Name)
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
            if (objFSO.FolderExists(configsourcepath & "\" & mo & "_" & da)) then
                ' If no images in this folder, use the default folder
                Set temp = FSO.GetFolder(configsourcepath & "\" & mo & "_" & da)
                if (temp.Files.Count > 0) then
                    Set SourceFolder = FSO.GetFolder(configsourcepath & "\" & mo & "_" & da)
                    folderPath = configcontents(1) & "\" & mo & "_" & da
                    moday = mo & "_" & da
                end if
            end if

        
        '-------------------------
        ' DEBUGGING: Tell user folder exists for this month
        '        WScript.Echo "You are using the following folder: " & SourceFolder.Path
        '-------------------------

        '-------------------------
        ' Select a random picture to use for the wallpaper
        '-------------------------
            max = SourceFolder.Files.Count
            min = 1
            Randomize
            therand = Int((max-min+1) * Rnd+min)

            temp = ""
            i = 0
            
            ' configkeepimages_orig = configkeepimages
            
            if (max<CInt(configkeepimages)) then
               ' WScript.echo "Changing max number of images from " & CInt(configkeepimages) & " to " & max
               configkeepimages = max
               ' If you want to keep 20 files in your destination folder, but only have 2 in your source folder,
               ' this sets you up to only keep as many files as are available.
               ' If you have only a few available files to choose from, this script
               ' will make multiple copies of them and add them to the destination
               ' folder. With this configuration check, it prevents the system from
               ' making excessive copies of your images.
            end if


            ' Select a file with qualifying extension
            For each file in SourceFolder.Files
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
                SelectNewWallpaper = defFile
            else
                SelectNewWallpaper = temp
            end if

            'WScript.echo "You have chosen: " &chr(13) & selectedwallpaper

    else
        '-------------------------
        ' If no configuration file has been found, but this subroutine has been called...exit application
        '-------------------------
        WScript.Quit
        ' Application has failed.
    end if
    


'-------------------------END OF SELECT NEW WALLPAPER FUNCTION-------------------------
End Function





Function CopyNewWallpaper()
    ' This function changes the registry settings to select the new wallpaper and other user preferences from the 
    '   configuration settings file.

    On Error Resume Next
    Err.Clear

    if (isnull(selectedwallpaper) OR selectedwallpaper = "") then
        msgbox("No wallpaper found in " & folderPath & "\" & selectedwallpaper)
    else
        Set DestFolder = FSO.GetFolder(configdestpath)

        ' Check to see if this image is already in the destiantion folder. If so, select a new image
        j = 1
        
        While (j<(CInt(configkeepimages)+1))
            'WScript.echo "Checking for duplicates of " & selectedwallpaper & "." & chr(13) & "Iteration " & j & "."
            For each file in DestFolder.Files
                temp1 = right(file.Name,len(file.Name)-instr(file.Name,"-"))
                if ((lcase(temp1)=lcase(selectedwallpaper) OR lcase(right(temp1,len(temp1)-12))=lcase(selectedwallpaper)) AND errorcount<=CInt(configkeepimages)) then
                   error="Duplicate"
                end if
            Next
            
            if (error="Duplicate") then
               logText = logText & chr(13) & "[DUPLICATE IMAGE     ]    [" & timestamp & "]    Found Duplicate File (" & j & ") - " & _
                           selectedwallpaper & ". Selecting a new wallpaper."
               'WScript.Echo "Duplicate found: " & selectedwallpaper & ". Calling SelectNewWallpaper Function."
               selectedwallpaper=SelectNewWallpaper
            else
                j=(CInt(configkeepimages)+1)
            end if
            error=""
            j=j+1
        Wend
        
        
        ' WScript.echo logText
        
        ' Delete old #1 wallpaper and all wallpapers over the keep limit
        For each file in DestFolder.Files
            temp1=CInt(left(file.Name,instr(file.Name,"-")-1))
            if (temp1<1) then
               temp1=1
            end if
            if ((temp1+1)<10 AND temp1>0) then
               temp2 = "0" & (temp1+1)
            else
                temp2 = "" & (temp1+1)
            end if
            'WScript.echo file.Name & chr(13) & "Temp1="&temp1 & chr(13) & "Temp2=" & temp2
            'WScript.echo temp1 & chr(13) & temp2 & chr(13) & right(file.Name,len(file.Name)-2) & _
            '  chr(13) & "will be moved from " & chr(13) &_
            '  configdestpath & "\" & file.Name & chr(13) & "to" &chr(13) &_
            '  configdestpath & "\" & temp2 & (right(file.Name,len(file.Name)-2))
            NewFileName = temp2 & "-" & (right(file.Name,len(file.Name)-instr(file.Name,"-")))
            if (objFSO.FileExists(configdestpath & "\" & NewFileName)) then
               objFSO.MoveFile configdestpath & "\" & NewFileName, configdestpath & "\" & _
                 temp2 & "-" & timestamp & NewFileName
            end if
            
            objFSO.MoveFile configdestpath & "\" & file.Name, configdestpath & "\" & _
                 temp2 & "-" & (right(file.Name,len(file.Name)-instr(file.Name,"-")))
            
        Next

        For each file in DestFolder.Files
            temp1=CInt(left(file.Name,instr(file.Name,"-")-1))
            if (temp1<2 OR temp1>CInt(configkeepimages)) then
               objFSO.DeleteFile file
            end if
            
            'if (left(temp.Name,2)
        Next
        
        ' Copy new wallpaper file over to the new location
        objFSO.CopyFile SourceFolder.Path & "\" & selectedwallpaper , DestFolder.Path & "\" & "01-" & selectedwallpaper
        
        ' Write the new information in the configuration file
        Set objWallFile = objFSO.CreateTextFile (scriptDirectory & wallFile, ForWriting)
        objWallFile.WriteLine("Get New Images From This Directory:")
        objWallFile.WriteLine(configsourcepath)
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Store Current Wallpaper Images in This Directory:")
        objWallFile.WriteLine(configdestpath)
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Keep This Number of Wallpaper Images:")
        objWallFile.WriteLine(CInt(configkeepimages_orig))
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Wallpaper Last Changed:")
        objWallFile.WriteLine(Now())

        ' Call another function to delete older wallpapers from the destination location
        

        

    end if

End Function






