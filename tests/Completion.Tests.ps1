$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'app\Core.ps1')
. (Join-Path $root 'app\State.ps1')
$count=0
function Assert($ok,$message){if(-not $ok){throw ('FAILED: '+$message)};$script:count++}
$dir=Join-Path ([IO.Path]::GetTempPath()) ('MzrbCompletion-'+[guid]::NewGuid().ToString('N'))
try{
 foreach($noteCount in @(0,1,2)){
  $warnings=New-Object Collections.Generic.List[string];$warnings.Add('Проверка <источника> & периода')
  $s=[pscustomobject]@{Automatic=$true;BasisPath='basis.xls';BasisSheet='TDSheet';Source='ledger.xlsx';Sheet='01.11.26';PeriodKey='2026-10';Warnings=$warnings}
  $notes=New-Object Collections.Generic.List[object]
  for($i=0;$i -lt $noteCount;$i++){$notes.Add([pscustomobject]@{SourceRow=5+$i;Contract='D-'+$i;Name='Drug';Reason='Not unique <&>';Relevant=($i -eq 0)})}
  $changes=New-Object Collections.Generic.List[object]
  $changes.Add([pscustomobject]@{BasisRow=9;Contract='D-1';Name='Drug <A>';Field='Commission';Before=[decimal]5;After=[decimal]2})
  $changes.Add([pscustomobject]@{BasisRow=10;Contract='D-2';Name='';Field='Name';Before='';After='Drug & B'})
  $plan=[pscustomobject]@{Notes=$notes;Changes=$changes;Skipped=@();Total=[decimal]12;PositionCount=2}
  $output=Join-Path $dir ('report-'+$noteCount+'.xls')
  # Exercise the same production function called after Excel.SaveAs, without requiring Excel.
  $completionOutput=@(Complete-MergedReport $s $plan $output)
  Assert ($completionOutput.Count -eq 0) 'Completion must not leak objects into the background job result'
  Assert ($s.ResultTotal -eq 12 -and $s.ResultPositions -eq 2 -and $s.ResultChanges -eq 2) 'Result metadata preserved'
  Assert ($s.AutomationNotes -is [object[]] -and $s.AutomationNotes.Count -eq $noteCount) 'Notes must be stable arrays, including zero and one item'
  $expectedWarnings=1+$(if($noteCount -gt 0){1}else{0})
  Assert ($s.ResultNotes -eq $expectedWarnings) 'Count relevant notes and source warnings exactly once'
  $audit=Join-Path $dir ('Проверки\report-'+$noteCount+'.html')
  Assert (Test-Path -LiteralPath $audit) 'Audit created after export'
  $text=[IO.File]::ReadAllText($audit)
  Assert ($text.Contains('Drug &lt;A&gt;') -and $text.Contains('Drug &amp; B') -and $text.Contains('12.00')) 'Audit contains escaped rows, completed names and total'
  Assert ($text.Contains('&lt;источника&gt; &amp; периода')) 'Source warning is encoded in audit'
  if($noteCount){Assert ($text.Contains('Not unique &lt;&amp;&gt;')) 'Matching reason is encoded in audit'}
  $notes.Add([pscustomobject]@{Relevant=$false})
  Assert ($s.AutomationNotes.Count -eq $noteCount) 'Session notes are a snapshot, not the mutable input list'
 }
}finally{if(Test-Path -LiteralPath $dir){Remove-Item -LiteralPath $dir -Recurse -Force}}
"$count completion checks passed"
