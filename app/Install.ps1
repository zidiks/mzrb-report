param([string]$Destination=(Join-Path $env:LOCALAPPDATA 'MzrbReport'))
$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework
try{
 . (Join-Path $PSScriptRoot 'Update.ps1')
 $Destination=[IO.Path]::GetFullPath($Destination)
 if($Destination -ne $PSScriptRoot){
  if(Test-Path -LiteralPath (Join-Path $Destination 'App.ps1')){throw 'Программа уже установлена. Откройте её и нажмите «Проверить обновления».'}
  [void][IO.Directory]::CreateDirectory($Destination)
  foreach($name in $script:PackageFiles){$path=Join-Path $Destination $name;[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path));[IO.File]::Copy((Join-Path $PSScriptRoot $name),$path,$true)}
 }
 $shell=New-Object -ComObject WScript.Shell;$link=$shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'Отчёт МЗ РБ.lnk'))
 $link.TargetPath=Join-Path $PSHOME 'powershell.exe';$link.Arguments='-NoLogo -NoProfile -STA -File "'+(Join-Path $Destination 'App.ps1')+'"';$link.WorkingDirectory=$Destination;$link.WindowStyle=7;$link.IconLocation=Join-Path $Destination 'Data\Report.ico';$link.Save()
 Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList @('-NoLogo','-NoProfile','-STA','-File',('"'+(Join-Path $Destination 'App.ps1')+'"')) -WorkingDirectory $Destination -WindowStyle Hidden | Out-Null
}catch{[void][Windows.MessageBox]::Show($_.Exception.Message,'Не удалось установить программу','OK','Error');exit 1}
