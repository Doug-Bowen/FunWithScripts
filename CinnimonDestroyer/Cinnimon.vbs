' WallpaperChanger.vbs - by Winkie - Last changed: 20100927
Option Explicit
 
Dim oShell, oScripting, aFilesArray
Dim sWallpaperDir, sFolder, sFile
Dim i , iFiles, iSelection
 
sFolder = "Destroyers\" ' Include quotes and last backslash
 
Set oShell = CreateObject("WScript.Shell")
Set oScripting = CreateObject("Scripting.FileSystemObject")
 
Set sFolder = oScripting.GetFolder(sFolder)
 
iFiles = CInt(sFolder.Files.Count)
 
ReDim aFilesArray(0)
For Each sFile in sFolder.Files
   ReDim Preserve aFilesArray(UBound(aFilesArray) + 1)
   aFilesArray(UBound(aFilesArray)) = sFile.ShortPath
Next
 
i = 1
Randomize
iSelection = Int((UBound(aFilesArray) - i + 1) * Rnd + i)
 
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", aFilesArray(iSelection)
oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, False