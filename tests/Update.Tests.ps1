$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
. (Join-Path $repo 'app\Update.ps1')
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('MzrbTests-'+[guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($testRoot)
$script:Assertions=0
function Assert($condition,[string]$message){if(-not $condition){throw ('FAILED: '+$message)};$script:Assertions++}
function Reject([scriptblock]$body,[string]$message){$rejected=$false;try{& $body | Out-Null}catch{$rejected=$true};Assert $rejected $message}
try{
 $zip=Join-Path $repo 'dist\mzrb-report.zip';$hash=Update-Hash $zip;$version=(App-Version (Join-Path $repo 'app')).ToString()
 $stage=Expand-Update $zip (Join-Path $testRoot 'stage') $hash $version
 Assert ((App-Version $stage).ToString() -eq $version) 'Valid package extracted'
 Reject {Expand-Update $zip (Join-Path $testRoot 'badHash') ('0'*64) $version} 'Checksum mismatch rejected'
 Assert (-not (Test-Path -LiteralPath (Join-Path $testRoot 'badHash'))) 'No writes before checksum verification'
 Reject {Expand-Update $zip (Join-Path $testRoot 'badVersion') $hash '99.0.0'} 'Wrong version rejected'
 foreach($url in @('http://github.com/zidiks/mzrb-report/releases/test','https://example.com/zidiks/mzrb-report/a','https://github.com/other/mzrb-report/a','https://github.com/zidiks/mzrb-report-evil/a')){Reject {GitHub-Uri $url 'github.com'} 'Untrusted release address rejected'}
 Reject {GitHub-Uri 'https://api.github.com/repos/other/repo/releases/assets/1' 'api.github.com'} 'Foreign repository API rejected'
 Reject {GitHub-Uri 'https://api.github.com/repos/zidiks/mzrb-report/releases/assets/1?redirect=evil' 'api.github.com'} 'Unexpected API query rejected'
 Assert ((GitHub-Uri 'https://api.github.com/repos/zidiks/mzrb-report/releases/assets/1' 'api.github.com').Host -eq 'api.github.com') 'Exact release asset API accepted'
 $evil=Join-Path $testRoot 'evil.zip';$archive=[IO.Compression.ZipFile]::Open($evil,[IO.Compression.ZipArchiveMode]::Create)
 try{[void]$archive.CreateEntry('../outside.ps1')}finally{$archive.Dispose()}
 Reject {Expand-Update $evil (Join-Path $testRoot 'evil') (Update-Hash $evil) $version} 'Zip traversal rejected'
 Assert (-not (Test-Path -LiteralPath (Join-Path $testRoot 'outside.ps1'))) 'No path traversal writes'
 $private=Join-Path $testRoot 'private.zip';$archive=[IO.Compression.ZipFile]::Open($private,[IO.Compression.ZipArchiveMode]::Create)
 try{[void]$archive.CreateEntry('Data/October.json')}finally{$archive.Dispose()}
 Reject {Expand-Update $private (Join-Path $testRoot 'private') (Update-Hash $private) $version} 'Personal data cannot be replaced by package'
 $target=Join-Path $testRoot 'installed';[void][IO.Directory]::CreateDirectory((Join-Path $target 'Data\Work'))
 [IO.File]::WriteAllText((Join-Path $target 'App.ps1'),'old app')
 [IO.File]::WriteAllText((Join-Path $target 'Core.ps1'),'old core')
 [IO.File]::WriteAllText((Join-Path $target 'State.ps1'),'old state')
 foreach($name in @('Settings.json','Names.json','October.json','Work\session.json')){[IO.File]::WriteAllText((Join-Path $target ('Data\'+$name)),('private '+$name))}
 # Failure during copying after two files changed must restore the old app.
 $lock=[IO.File]::Open((Join-Path $target 'State.ps1'),[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
 try{Reject {Apply-Package $stage $target} 'Locked destination rejected'}finally{$lock.Dispose()}
 Assert ([IO.File]::ReadAllText((Join-Path $target 'App.ps1')) -eq 'old app') 'Rollback restored earlier replaced app'
 Assert ([IO.File]::ReadAllText((Join-Path $target 'Core.ps1')) -eq 'old core') 'Rollback restored earlier replaced core'
 $backup=Apply-Package $stage $target
 Assert (Test-Path -LiteralPath (Join-Path $backup 'App.ps1')) 'Backup retained'
 Assert ((App-Version $target).ToString() -eq $version) 'Installed version updated'
 foreach($name in @('Settings.json','Names.json','October.json','Work\session.json')){Assert ([IO.File]::ReadAllText((Join-Path $target ('Data\'+$name))) -eq ('private '+$name)) 'Private file preserved'}
 # A fresh public October install must work without a personal calibration file.
 Assert (-not (Test-Path -LiteralPath (Join-Path $stage 'Data\October.json'))) 'Release contains no personal profile'
 Write-Output ("$script:Assertions updater checks passed")
}finally{
 # Only the explicitly created, resolved test directory may be deleted.
 $resolved=[IO.Path]::GetFullPath($testRoot);$temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
 if($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('MzrbTests-')){Remove-Item -LiteralPath $resolved -Recurse -Force}
}
