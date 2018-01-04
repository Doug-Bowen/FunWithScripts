Dim fso, folder, files, WshShell, Return, FileCount, FileList(), FileNum, RandFile, PlayList(2), Played
  
'Get the folder info
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("Cosmonauts\")

    Randomize 'Init random number generator
    Set files = folder.Files
    ReDim FileList(CInt(files.Count - 1), 2)
    FileCount = 0
    for each folderIdx in files
        'load the file path/name
        FileList(FileCount, 0) = folderIdx.ParentFolder
        'load the file type
        FileList(FileCount, 1) = folderIdx.Type
        'load the short path
        FileList(FileCount, 2) = folderIdx.Name
        FileCount = FileCount + 1
        'msgbox(folderIdx.Type)
    next
    
    'Pick a random file
    RandFile = Int(((files.Count - 1) * Rnd) + 1)
    'msgbox(FileList(RandFile, 0) & " " & FileList(RandFile, 2))
    Set CmdExec = WshShell.Exec("mplayer.exe -noidle -really-quiet -fs " & """" & FileList(RandFile, 0) & "\" & FileList(RandFile, 2) & """")

