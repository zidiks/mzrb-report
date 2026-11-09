$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'app\State.ps1')
$count=0
function Assert($ok,$message){if(-not $ok){throw ('FAILED: '+$message)};$script:count++}
$shell=[powershell]::Create()
try{
 [void]$shell.AddScript('function Original-Worker { throw "Format failure at F9:F12" }; Original-Worker')
 $handle=$shell.BeginInvoke();$wrapped=$null
 try{[void]$shell.EndInvoke($handle)}catch{$wrapped=$_}
 Assert ($null -ne $wrapped) 'Fixture fails through a real asynchronous PowerShell invocation'
 $record=Resolve-WorkError $wrapped
 Assert ($record.ScriptStackTrace.Contains('Original-Worker')) 'Recover original worker stack through EndInvoke wrappers'
 Assert ($record.ToString().Contains('F9:F12')) 'Preserve range context from worker failure'
 $details=Work-ErrorDetails $wrapped
 Assert ($details.Contains('Original-Worker') -and $details.Contains('throw') -and $details.Contains('HRESULT=0x')) 'Diagnostic text contains original stack, source position and error code'
}finally{$shell.Dispose()}
try{throw 'Direct error'}catch{$direct=$_}
Assert ([object]::ReferenceEquals((Resolve-WorkError $direct),$direct)) 'Direct errors retain their original record'
"$count worker error checks passed"
