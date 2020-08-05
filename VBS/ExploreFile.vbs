Set exec = CreateObject("WScript.Shell").Exec("powershell.exe -sta  -Command Add-Type -Assembly System.Windows.Forms; [System.Windows.Forms.Clipboard]::GetText()")
exec.StdIn.Close
CreateObject("WScript.Shell").Exec("explorer.exe /select," + exec.StdOut.ReadAll)
-----------------------------7e42cb1f10304--
