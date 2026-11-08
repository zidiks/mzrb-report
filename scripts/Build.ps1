param([string]$Tag,[string]$OutputDirectory=(Join-Path (Split-Path $PSScriptRoot -Parent) 'dist'))
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent;$app=Join-Path $root 'app'
. (Join-Path $app 'Update.ps1')
$version=(App-Version $app).ToString()
if($Tag -and $Tag -cne ('v'+$version)){throw 'Git tag does not match app/Version.json'}
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$zipPath=Join-Path $OutputDirectory 'mzrb-report.zip'
if(Test-Path -LiteralPath $zipPath){[IO.File]::Delete($zipPath)}
$zip=[IO.Compression.ZipFile]::Open($zipPath,[IO.Compression.ZipArchiveMode]::Create)
try{
 foreach($name in $script:PackageFiles){
  $file=Join-Path $app $name
  if(-not (Test-Path -LiteralPath $file -PathType Leaf)){throw ('Missing release file: '+$name)}
  [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip,$file,$name,[IO.Compression.CompressionLevel]::Optimal)
 }
}finally{$zip.Dispose()}
$hash=Update-Hash $zipPath
[IO.File]::WriteAllText((Join-Path $OutputDirectory 'SHA256SUMS.txt'),($hash+'  mzrb-report.zip'+"`n"),(New-Object Text.UTF8Encoding($false)))
Write-Output ('Built '+$version+': '+$zipPath)
