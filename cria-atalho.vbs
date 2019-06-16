strAppPath = "C:\Users\cnnr\Documents\local-arquivo\arquivo.txt"

Set objShell = CreateObject("WScript.Shell")
objDesktop = objShell.SpecialFolders("Desktop")
Set objLink = objShell.CreateShortcut(objDesktop & "\arquivo-atalho.lnk")

objLink.TargetPath = strAppPath
objLink.WindowStyle = 3
objLink.Save

WScript.Quit