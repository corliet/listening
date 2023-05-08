function install {
    $target = Join-Path -Path $listeningDir -ChildPath "listening.html"
    $shortcutPath = Join-Path -Path $home -ChildPath "Desktop\リスニング.lnk"
    $WShell = New-Object -ComObject WScript.Shell
    $shortcut = $WShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetPath
    $shortcut.Save()
}
