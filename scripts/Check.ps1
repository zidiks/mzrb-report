$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
foreach($file in Get-ChildItem -LiteralPath $root -Filter '*.ps1' -Recurse){
 if($file.FullName.StartsWith((Join-Path $root 'dist'))){continue}
 $tokens=$null;$errors=$null;[void][Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors)
 if($errors.Count){throw ($file.Name+': '+($errors.Message -join '; '))}
 if($file.DirectoryName -eq (Join-Path $root 'app') -and $file.Name -ne 'Version.json'){
  $bytes=[IO.File]::ReadAllBytes($file.FullName)
  if($bytes.Length -lt 3 -or $bytes[0] -ne 239 -or $bytes[1] -ne 187 -or $bytes[2] -ne 191){throw ('Windows PowerShell requires UTF-8 BOM: '+$file.Name)}
 }
}
Add-Type -AssemblyName PresentationFramework
foreach($name in @('Main.xaml','Review.xaml')){
 $reader=New-Object Xml.XmlNodeReader ([xml][IO.File]::ReadAllText((Join-Path $root ('app\'+$name))))
 try{[void][Windows.Markup.XamlReader]::Load($reader)}finally{$reader.Close()}
}
& (Join-Path $root 'tests\Core.Tests.ps1')
& (Join-Path $root 'tests\Layout.Tests.ps1')
& (Join-Path $root 'tests\Excel.Template.Tests.ps1')
& (Join-Path $root 'tests\Window.Tests.ps1')
& (Join-Path $PSScriptRoot 'Build.ps1')
& (Join-Path $root 'tests\Update.Tests.ps1')
Write-Output 'All checks passed'
