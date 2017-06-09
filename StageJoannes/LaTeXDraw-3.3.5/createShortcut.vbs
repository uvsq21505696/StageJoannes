Set shell = WScript.CreateObject("WScript.Shell")
strProg = shell.SpecialFolders("Programs")
Set link= shell.CreateShortcut(strProg+"\LaTeXDraw.LNK")
link.TargetPath = "C:\Program Files\latexdraw\LaTeXDraw.jar"
link.Save
