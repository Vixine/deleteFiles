Option Explicit
Dim username, desktop, desktopfolder, shortcuts, output
Dim scrShell, objFSO, file

desktop = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\"
output = ""

desktop = "C:\Users\Yisroel\OneDrive\Documents\GitHub\deleteFiles"
Set objFSO = CreateObject("Scripting.FileSystemObject")
set desktopfolder = objFSO.GetFolder(desktop)
set shortcuts = desktopfolder.Files
for each file in shortcuts
    if file.type="Shortcut" OR file.type="Remote Desktop Connection" Then
        output = output & file.Path & vbCrLf
        'objFSO.DeleteFile(file.Path)
    else
        
    end if
    
next
WScript.Echo "Files to Delete:" & vbCrLf & output
