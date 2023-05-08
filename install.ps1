function install {
    $listeningDir = Join-Path -Path $HOME -ChildPath "リスニング"
    $target = Join-Path -Path $listeningDir -ChildPath "listening.html"
    $shortcutPath = Join-Path -Path $HOME -ChildPath "Desktop\リスニング.lnk"
    if (Test-Path -Path $shortcutPath) {
        Remove-Item -Path $shortcutPath
    }
    $WShell = New-Object -ComObject WScript.Shell
    $shortcut = $WShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = [String]$target
    $shortcut.Save()
    Write-Host "Successfully created shortcut."
}

install
