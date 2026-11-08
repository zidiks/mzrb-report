param([Parameter(Mandatory=$true)][string]$Request)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Update.ps1')
Add-Type -AssemblyName PresentationFramework
$target=$null;$updated=$false
try{
 $r=[IO.File]::ReadAllText($Request) | ConvertFrom-Json;$target=[IO.Path]::GetFullPath($r.Target)
 $owner=Get-Process -Id ([int]$r.Owner) -ErrorAction SilentlyContinue
 if($owner -and $owner.StartTime.ToUniversalTime().Ticks -eq $r.OwnerStarted){if(-not $owner.WaitForExit(180000)){throw 'Программа ещё открыта. Обновление отменено.'}}
 $stage=Expand-Update $r.Zip (Join-Path (Split-Path $Request -Parent) 'verified') $r.Sha256 $r.Version
 [void](Apply-Package $stage $target);$updated=$true
}catch{
 if($target -and (Test-Path -LiteralPath (Join-Path $target 'Data'))){[IO.File]::WriteAllText((Join-Path $target 'Data\Ошибка обновления.txt'),$_.ToString(),[Text.Encoding]::UTF8)}
 [void][Windows.MessageBox]::Show(('Не удалось установить обновление: '+$_.Exception.Message),'Отчёт МЗ РБ','OK','Warning')
}
if($target -and (Test-Path -LiteralPath (Join-Path $target 'App.ps1'))){
 $args=@('-NoLogo','-NoProfile','-STA','-File',('"'+(Join-Path $target 'App.ps1')+'"'))
 if($updated){$args+='-UpdatedFrom';$args+='GitHub'}
 Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $args -WorkingDirectory $target -WindowStyle Hidden | Out-Null
}
