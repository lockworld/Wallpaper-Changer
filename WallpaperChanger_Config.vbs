' VBScript File

 Option Explicit

 'WScript.Echo "WallpaperChanger2_Config.vbs has been called"



Dim _
configcontents(), _
configexists, _
configfilepath, _
configimage, _
configposition, _
configslideshow, _
currentImage, _
defFile, _
explines, _
extName, _
file, _
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
MyFolder, _
MyFiles, _
objFile, _
objFolder, _
objFSO, _
objLogFile, _
objNet, _
objReadFile, _
objShell, _
objStream, _
objWallFile, _
ofolder, _
oldFilePath, _
oSHApp, _
scriptPath, _
SPath, _
strComputer, _
strDesktop, _
sUserName, _
sWinDir, _
sWallPaper, _
SysFolder, _
temp, _
uinput, _
userPath, _
userFile, _
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



    ' Assign the file open/read variables (That won't be changed later in the program)
    ForAppending = 8
    ForReading = 1
    ForWriting = 2 'ForWriting will delete the existing contents before writing to the file

    ' Find the path the the current WallpaperChanger script.
    scriptPath = Left(WScript.ScriptFullName, InstrRev(WScript.ScriptFullName, WScript.ScriptName) -1)
    

    ' This is the path where the configuration settings file will be found (or created).
    wallDirectory = scriptPath
   

    ' Assigns the filename and path to search for or create the configuration settings file
    wallFile = "WallpaperChanger Settings.txt"



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
    if (explines <> foundlines) then
        ' If size indicates an error in wallFile, call CreateConfigurationSettingsFile function to recreate it
        WScript.Echo "An error has occured with the configuration file: " & vbNewLine & wallDirectory & wallFile _
        & vbNewLine & "Executing built-in pause for 5 seconds to rebuild file."
        WScript.Sleep(2500)

        CreateConfigurationSettingsFile

    end if
    
    if (explines = foundlines) then
        ' If the file passes verification, read the file
        ReadWallFile

    else
        '-------------------------
        ' OPTION: Uncomment the line below to include a failure message before quitting the application
        '-------------------------
            WScript.Echo "The application was not able to read or create the configuration file. Please try again later."
        
        ' If after two tries, file doesn't verify, quit silently
        WScript.Quit

        

    end if



  ' Prompt user for changes
  GetUserInput
  
  
  temp = """" & scriptPath & "\WallpaperChanger.vbs"""
  objShell.Run(temp)
  
  
  objWallFile.close
  Set objWallFile = Nothing
    
  Set objFSO = Nothing
  Set objNet = Nothing
  Set objShell = Nothing
  Set oSHApp = Nothing
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

        
    ' Close the file
    objWallFile.close
    Set objWallFile = Nothing


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
            
            ' Close the file
            objWallFile.close
            Set objWallFile = Nothing

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
        configimage = configcontents(3)
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
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Wallpaper Position:")
    objWallFile.WriteLine("2")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Include My Pictures Slideshow?")
    objWallFile.WriteLine("No")
    objWallFile.WriteLine("")
    objWallFile.WriteLine("Wallpaper Last Changed:")
    objWallFile.WriteLine("Never")
    
    ' Close the file
    objWallFile.close
    Set objWallFile = Nothing

    '-------------------------
    ' DEBUGGING:
        WScript.Echo "New configuration file successfully created."
    '-------------------------
    
    WScript.Sleep 10000

    GetSetConfigFile

    VerifyConfigSettingsFileLines

    

    'GetUserInput
    ' Calls a separate script to allow the user to review and modify the configuration settings for the program.

    

End Function







Function GetUserInput()
    
    ' This function uses Internet Explorer to
    ' create a dialog and prompt for user input.
    '
    ' Version:             2.10
    ' Last modified:       2010-09-28
    '
    ' Argument:   [string] prompt text, e.g. "Please enter your name:"
    ' Returns:    [string] the user input typed in the dialog screen
    '
    ' Written by Rob van der Woude
    ' http://www.robvanderwoude.com
    ' Error handling code written by Denis St-Pierre
        Dim objIE

        ' Create an IE object
        Set objIE = CreateObject( "InternetExplorer.Application" )

        ' Specify some of the IE window's settings
        objIE.Navigate "about:blank"
        objIE.Document.Title = "Wallpaper Changer Configuration" ' " " & String( 100, "." )
        objIE.ToolBar        = False
        objIE.Resizable      = true
        objIE.StatusBar      = False
        objIE.Width          = 700
        objIE.Height         = 500

        ' Center the dialog window on the screen
        With objIE.Document.ParentWindow.Screen
            objIE.Left = (.AvailWidth  - objIE.Width ) \ 2
            objIE.Top  = (.Availheight - objIE.Height) \ 2
        End With


        ' Precompile combo boxes to use correctly selected data
        dim ttov1, ttov2, ttov3
        ttov1 = ""
        ttov2 = ""
        ttov3 = ""
        if (configposition = "0") then
            ttov1 = " selected"
        end if
        if (configposition = "1") then
            ttov2 = " selected"
        end if
        if (configposition = "2") then
            ttov3 = " selected"
        end if

        dim ss
        ss = ""

        if (configslideshow = "Yes") then
            ss = " selected"
        end if


        ' Wait till IE is ready
        Do While objIE.Busy
            WScript.Sleep 200
        Loop
        ' Insert the HTML code to prompt for user input
        objIE.Document.Body.InnerHTML = "<div align=""left""><h4>Custom Wallpaper Configuration Settings:</h4>" & vbCrLf _
                                      & "<p><b>Enter the path to your wallpapers folder: </b><br/><input type=""text"" size=""20"" " _
                                      & "id=""UserPath"" value=""" & configfilepath & """></p>" & vbCrLf _
                                      & "<p><b>Select how you want your wallpaper to appear: </b><br/>" _
                                      & "<select id=""TileType"" value=""" & configposition & """>" & vbCrLf _
                                      & "  <option value=""0""" & ttov1 & ">Center</option>" & vbCrLf _
                                      & "  <option value=""1""" & ttov2 & ">Tile</option>" & vbCrLf _
                                      & "  <option value=""2""" & ttov3 & ">Stretch</option>" & vbCrLf _
                                      & "</select>" & vbCrLf _
                                      & "<p><b>Use the same directory for the ""My Pictures Slideshow"" screensaver?</b><br/>" & vbCrLf _
                                      & "<select id=""Slideshow"" value=""" & configslideshow & """>" & vbCrLf _
                                      & "  <option value=""No"">No</option>" & vbCrLf _
                                      & "  <option value=""Yes""" & ss & ">Yes</option" & vbCrLf _
                                      & "</select></p>" & vbCrLf _
                                      & "<p><input type=""hidden"" id=""OK"" " _
                                      & "name=""OK"" value=""0""><br/>" _
                                      & "<input type=""submit"" value="" OK "" " _
                                      & "OnClick=""VBScript:OK.Value=1""></p></div>"
        ' Hide the scrollbars
        objIE.Document.Body.Style.overflow = "auto"
        ' Make the window visible
        objIE.Visible = True
        ' Set focus on input field
        objIE.Document.All.UserPath.Focus

        ' Wait till the OK button has been clicked
        On Error Resume Next
        Do While objIE.Document.All.OK.Value = 0 
            WScript.Sleep 200
            ' Error handling code by Denis St-Pierre
            If Err Then ' user clicked red X (or alt-F4) to close IE window
                IELogin = Array( "", "" )
                objIE.Quit
                Set objIE = Nothing
                Exit Function
            End if
        Loop
        On Error Goto 0


        ' Read the user input from the dialog window
        ' and save it to the settings file

        Set objWallFile = objFSO.CreateTextFile (wallDirectory & wallFile, True)

    
        objWallFile.WriteLine("Wallpaper Directory:")
        objWallFile.WriteLine(objIE.Document.All.UserPath.Value)
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Current Wallpaper:")
        objWallFile.WriteLine(configcontents(4))
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Wallpaper Position:")
        objWallFile.WriteLine(objIE.Document.All.TileType.Value)
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Include 'My Pictures Slideshow?'")
        objWallFile.WriteLine(objIE.Document.All.Slideshow.Value)
        objWallFile.WriteLine("")
        objWallFile.WriteLine("Wallpaper Last Changed:")
        objWallFile.WriteLine(Now())

    
        ' Close and release the object
        objIE.Quit
        Set objIE = Nothing

        MsgBox("Your settings have been saved!")
End Function


