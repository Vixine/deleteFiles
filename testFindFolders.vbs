Dim fso, folder, files, OutputFile
Dim strPath, username, desktop

' Create a FileSystemObject  
Set fso = CreateObject("Scripting.FileSystemObject")

' Define folder we want to list files from
username = CreateObject("WScript.Network").UserName
desktop = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & "\OneDrive\Desktop\"
WScript.Echo desktop
'desktop = "C:\users\" & username & "\Desktop\"
strPath = desktop

Set folder = fso.GetFolder(strPath)
Set files = folder.Files

' Create text file to output test data
Set OutputFile = fso.CreateTextFile("OneScriptOutput.txt", True)

' Loop through each file  
For each item In files

  ' Output file properties to a text file
  OutputFile.WriteLine(item.Name)
  OutputFile.WriteLine(item.Attributes)
  OutputFile.WriteLine(item.DateCreated)
  OutputFile.WriteLine(item.DateLastAccessed)
  OutputFile.WriteLine(item.DateLastModified)
  OutputFile.WriteLine(item.Drive)
  OutputFile.WriteLine(item.Name)
  OutputFile.WriteLine(item.ParentFolder)
  OutputFile.WriteLine(item.Path)
  OutputFile.WriteLine(item.ShortName)
  OutputFile.WriteLine(item.ShortPath)
  OutputFile.WriteLine(item.Size)
  OutputFile.WriteLine(item.Type)   
  OutputFile.WriteLine("")
  
Next

' Close text file
OutputFile.Close
