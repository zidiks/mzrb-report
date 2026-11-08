# Windows PowerShell 5.1. Only the fixed public GitHub repository is trusted.
$script:ReleaseRepository='zidiks/mzrb-report'
$script:PackageFiles=@('App.ps1','Core.ps1','State.ps1','Update.ps1','ApplyUpdate.ps1','Install.ps1','Main.xaml','Review.xaml','Version.json','Запустить.cmd','Установить.cmd','Как пользоваться.txt','Data/Layout.xls','Data/Report.ico')
function App-Version([string]$root=$PSScriptRoot){
 $info=[IO.File]::ReadAllText((Join-Path $root 'Version.json')) | ConvertFrom-Json
 if($info.repository -ne $script:ReleaseRepository -or $info.version -notmatch '^\d+\.\d+\.\d+$'){throw 'Повреждена информация о версии программы.'}
 return [version]$info.version
}
function Update-Hash([string]$path){
 $stream=[IO.File]::OpenRead($path);$sha=[Security.Cryptography.SHA256]::Create()
 try{return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose();$stream.Dispose()}
}
function GitHub-Uri([string]$url,[string]$hostName){
 $uri=[uri]$url
 if(-not $uri.IsAbsoluteUri -or $uri.Scheme -ne 'https' -or $uri.Host -ne $hostName -or $uri.UserInfo -or $uri.Query -or $uri.Fragment){throw 'Недопустимый адрес обновления.'}
 if($hostName -eq 'api.github.com'){
  if($uri.AbsolutePath -notmatch ('^/repos/'+[regex]::Escape($script:ReleaseRepository)+'/releases/assets/\d+$')){throw 'Недопустимый адрес пакета GitHub.'}
 }elseif($hostName -ne 'github.com' -or -not $uri.AbsolutePath.StartsWith('/'+$script:ReleaseRepository+'/',[StringComparison]::Ordinal)){throw 'Недопустимый адрес обновления.'}
 return $uri
}
function Latest-Release([string]$root=$PSScriptRoot){
 [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
 $url='https://api.github.com/repos/'+$script:ReleaseRepository+'/releases/latest'
 $release=Invoke-RestMethod -Uri $url -Headers @{'User-Agent'='MzrbReport';Accept='application/vnd.github+json'} -TimeoutSec 25
 if($release.draft -or $release.prerelease -or $release.tag_name -notmatch '^v(\d+\.\d+\.\d+)$'){throw 'Не удалось определить стабильную версию.'}
 $version=[version]$matches[1];$current=App-Version $root
 if($version -le $current){return [pscustomobject]@{Available=$false;Version=$current.ToString()}}
 $assets=@($release.assets | Where-Object {$_.name -ceq 'mzrb-report.zip'})
 if($assets.Count -ne 1 -or $assets[0].digest -notmatch '^sha256:([a-fA-F0-9]{64})$'){throw 'У релиза нет проверяемого пакета обновления.'}
 $hash=$matches[1].ToLowerInvariant();$asset=$assets[0]
 [void](GitHub-Uri $asset.browser_download_url 'github.com')
 if($asset.id -le 0){throw 'Не найден идентификатор пакета GitHub.'}
 $apiUrl='https://api.github.com/repos/'+$script:ReleaseRepository+'/releases/assets/'+[string]$asset.id
 if($asset.size -le 0 -or $asset.size -gt 10MB){throw 'Некорректный размер пакета обновления.'}
 return [pscustomobject]@{Available=$true;Version=$version.ToString();Url=$asset.browser_download_url;ApiUrl=$apiUrl;Sha256=$hash;Size=[long]$asset.size}
}
function Expand-Update([string]$zipPath,[string]$destination,[string]$sha256,[string]$version){
 if($sha256 -notmatch '^[a-fA-F0-9]{64}$' -or (Update-Hash $zipPath) -cne $sha256.ToLowerInvariant()){throw 'Проверка целостности обновления не пройдена. Файлы программы не изменены.'}
 if(Test-Path -LiteralPath $destination){throw 'Папка распаковки уже существует.'}
 Add-Type -AssemblyName System.IO.Compression
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $archive=[IO.Compression.ZipFile]::OpenRead($zipPath)
 try{
  $seen=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal);$total=[long]0
  foreach($entry in $archive.Entries){
   if($entry.FullName -cnotin $script:PackageFiles -or -not $seen.Add($entry.FullName)){throw ('Недопустимый файл в обновлении: '+$entry.FullName)}
   $total+=$entry.Length;if($entry.Length -gt 10MB -or $total -gt 20MB){throw 'Пакет обновления слишком большой.'}
  }
  if($seen.Count -ne $script:PackageFiles.Count){throw 'В пакете обновления отсутствуют необходимые файлы.'}
  [void][IO.Directory]::CreateDirectory($destination)
  foreach($entry in $archive.Entries){$target=Join-Path $destination $entry.FullName;[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target));[IO.Compression.ZipFileExtensions]::ExtractToFile($entry,$target)}
  if((App-Version $destination).ToString() -cne $version){throw 'Версия пакета не совпадает с версией релиза.'}
  return $destination
 }finally{$archive.Dispose()}
}
function Download-Update($release){
 [void](GitHub-Uri $release.Url 'github.com')
 [void](GitHub-Uri $release.ApiUrl 'api.github.com')
 $base=Join-Path ([IO.Path]::GetTempPath()) ('MzrbReport-'+[guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($base)
 $zip=Join-Path $base 'package.zip'
 [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
 try{Invoke-WebRequest -UseBasicParsing -Uri $release.ApiUrl -Headers @{'User-Agent'='MzrbReport';Accept='application/octet-stream'} -OutFile $zip -TimeoutSec 45 -ErrorAction Stop}
 catch{Invoke-WebRequest -UseBasicParsing -Uri $release.Url -Headers @{'User-Agent'='MzrbReport'} -OutFile $zip -TimeoutSec 45 -ErrorAction Stop}
 if((Get-Item -LiteralPath $zip).Length -ne $release.Size){throw 'Пакет скачан не полностью. Повторите попытку.'}
 $stage=Expand-Update $zip (Join-Path $base 'app') $release.Sha256 $release.Version
 return [pscustomobject]@{Stage=$stage;Zip=$zip;Sha256=$release.Sha256;Version=$release.Version}
}
function Apply-Package([string]$stage,[string]$target){
 $stage=[IO.Path]::GetFullPath($stage);$target=[IO.Path]::GetFullPath($target)
 if($target -eq $stage -or -not (Test-Path -LiteralPath (Join-Path $target 'App.ps1'))){throw 'Не найдена установленная программа.'}
 [void](App-Version $stage)
 $backup=Join-Path $target ('Data\Updates\Backups\'+[datetime]::UtcNow.ToString('yyyyMMddHHmmss')+'-'+[guid]::NewGuid().ToString('N'))
 [void][IO.Directory]::CreateDirectory($backup)
 # Back up every replaced file before changing anything; personal Data files are outside this list.
 foreach($name in $script:PackageFiles){
  if(-not (Test-Path -LiteralPath (Join-Path $stage $name) -PathType Leaf)){throw ('Неполный пакет: '+$name)}
  $old=Join-Path $target $name
  if(Test-Path -LiteralPath $old){$copy=Join-Path $backup $name;[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($copy));[IO.File]::Copy($old,$copy,$false)}
 }
 $written=New-Object Collections.Generic.List[string]
 try{
  foreach($name in $script:PackageFiles){
   $dest=Join-Path $target $name;[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($dest))
   $temporary=$dest+'.'+[guid]::NewGuid().ToString('N')+'.tmp'
   try{
    [IO.File]::Copy((Join-Path $stage $name),$temporary,$false)
    if(Test-Path -LiteralPath $dest){[IO.File]::Replace($temporary,$dest,[NullString]::Value)}else{[IO.File]::Move($temporary,$dest)}
    $written.Add($name)
   }finally{if(Test-Path -LiteralPath $temporary){[IO.File]::Delete($temporary)}}
  }
  if((App-Version $target) -ne (App-Version $stage)){throw 'Обновление не завершилось.'}
 }catch{
  $failure=$_
  foreach($name in $written){$previous=Join-Path $backup $name;$dest=Join-Path $target $name;if(Test-Path -LiteralPath $previous){[IO.File]::Copy($previous,$dest,$true)}elseif(Test-Path -LiteralPath $dest){[IO.File]::Delete($dest)}}
  throw $failure
 }
 return $backup
}
function Start-AppUpdate($package,[string]$target,[int]$ownerProcess){
 $owner=Get-Process -Id $ownerProcess
 $request=Join-Path (Split-Path $package.Stage -Parent) 'request.json'
 [ordered]@{Target=$target;Zip=$package.Zip;Sha256=$package.Sha256;Version=$package.Version;Owner=$ownerProcess;OwnerStarted=$owner.StartTime.ToUniversalTime().Ticks} | ConvertTo-Json | Set-Content -LiteralPath $request -Encoding UTF8
 $helper=Join-Path $package.Stage 'ApplyUpdate.ps1'
 Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList @('-NoLogo','-NoProfile','-STA','-File',('"'+$helper+'"'),'-Request',('"'+$request+'"')) -WindowStyle Hidden | Out-Null
}
