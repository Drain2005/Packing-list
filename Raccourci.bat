@echo off
echo Creation du raccourci sur le Bureau...

powershell -command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%userprofile%\Desktop\Generateur de fichier.lnk'); $Shortcut.TargetPath = '%~dp0LANCER L''APPLICATION.bat'; $Shortcut.WorkingDirectory = '%~dp0'; $Shortcut.Description = 'Generateur de fichier '; $Shortcut.IconLocation = 'shell32.dll,100'; $Shortcut.Save()"

echo ✅ Raccourci cree sur le Bureau !
echo.
echo Vous pouvez maintenant double-cliquer sur :
echo    « Generateur DE FICHIER » sur votre Bureau
echo.
pause