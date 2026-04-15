$folder   = $PSScriptRoot
$batPath  = Join-Path $folder "run_quiz.bat"
$lnkPath  = Join-Path $folder "Launch SIE Quizzer.lnk"

$shell      = New-Object -ComObject WScript.Shell
$shortcut   = $shell.CreateShortcut($lnkPath)
$shortcut.TargetPath       = $batPath
$shortcut.WorkingDirectory = $folder
$shortcut.WindowStyle      = 1
$shortcut.Description      = "SIE Exam Quizzer"
$shortcut.Save()
Write-Host "Folder shortcut created: $lnkPath"
