$folder = $Env:USERPROFILE + "\Desktop\*"
Remove-Item -Path $folder -Include *.lnk, *.rdp

