function Save-Session($s,[string]$path){[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path));[IO.File]::WriteAllText($path,($s | ConvertTo-Json -Depth 30),(New-Object Text.UTF8Encoding($true)))}
function Restore-Session([string]$path){
 $s=Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
 if($s.Version -ne 4 -or $null -eq $s.Lines -or $null -eq $s.Events){throw 'Не удалось восстановить сохранённую проверку.'}
 foreach($field in @('Lines','Groups','Events','Expenses')){$items=New-Object Collections.Generic.List[object];foreach($item in $s.$field){$items.Add($item)};$s.$field=$items}
 $warnings=New-Object Collections.Generic.List[string];foreach($w in $s.Warnings){$warnings.Add([string]$w)};$s.Warnings=$warnings
 foreach($l in $s.Lines){foreach($f in $script:Fields){$l.$f=Money $l.$f}};foreach($ev in $s.Events){$ev.Amount=Money $ev.Amount}
 return $s
}
function Session-Current($s,[string]$source,[string]$sheet,[string]$period,[string]$template,[string]$basisPath='',[string]$basisSheet=''){
 if($null -eq $s -or $s.Source -ne $source -or $s.Sheet -ne $sheet -or $s.PeriodKey -ne $period -or $s.Template -ne $template){return $false}
 if($basisPath){if($s.BasisPath -ne $basisPath -or $s.BasisSheet -ne $basisSheet -or -not (Test-Path -LiteralPath $basisPath -PathType Leaf)){return $false};if($s.BasisFileHash -ne (File-Sha256 $basisPath)){return $false}}elseif($s.BasisPath){return $false}
 if(-not (Test-Path -LiteralPath $source -PathType Leaf)){return $false}
 return $s.SourceFileHash -eq (File-Sha256 $source) -and $s.TemplateFileHash -eq (File-Sha256 $template) -and $s.EngineHash -eq (File-Sha256 (Join-Path $PSScriptRoot 'Core.ps1'))
}
function Cache-Path($source,$sheet,$period,$basisPath='',$basisSheet=''){return Join-Path $PSScriptRoot ('Data\Work\'+(Get-Hash ($source.ToLowerInvariant()+'|'+$sheet+'|'+$period+'|'+$basisPath.ToLowerInvariant()+'|'+$basisSheet))+'.json')}
function Save-Cache($s){Save-Session $s (Cache-Path $s.Source $s.Sheet $s.PeriodKey ([string]$s.BasisPath) ([string]$s.BasisSheet))}
function Pending-Groups($s){$ids=@{};foreach($l in $s.Lines){if(-not $l.Reviewed){$ids[$l.Group]=$true}};return @($s.Groups | Where-Object {$ids.ContainsKey($_.Id)})}
function Needs-Review($s){if($s.Automatic){return $false};return @(Pending-Groups $s).Count -gt 0 -or ($s.Warnings.Count -gt 0 -and -not $s.Confirmed)}
function Unique-Output($s,[bool]$draft){
 $dir=Join-Path ([IO.Path]::GetDirectoryName($s.Source)) 'Готовые отчёты';[void][IO.Directory]::CreateDirectory($dir)
 $date=[datetime]::ParseExact($s.PeriodKey,'yyyy-MM',[Globalization.CultureInfo]::InvariantCulture);$stem=$date.ToString('MMyyyy')+$(if($draft){'_черновик'}else{'_отчёт'});$ext='.xls';$path=Join-Path $dir ($stem+$ext);$number=2
 while(Test-Path -LiteralPath $path){$path=Join-Path $dir ($stem+' ('+$number+')'+$ext);$number++}
 return $path
}
function Analyze-BuiltIn([string]$source,[string]$sheet,[datetime]$period,[bool]$fresh=$false){
 $template=Join-Path $PSScriptRoot 'Data\Layout.xls'
 if(-not (Test-Path -LiteralPath $template -PathType Leaf)){throw 'Не найден встроенный формат отчёта. Восстановите папку программы.'}
 $key=$period.ToString('yyyy-MM');$cache=Cache-Path $source $sheet $key
 if(-not $fresh -and (Test-Path -LiteralPath $cache)){
  try{$saved=Restore-Session $cache;if(Session-Current $saved $source $sheet $key $template){return $saved}}catch{}
 }
 $s=Analyze-Files $source $template $sheet $period $true
 $s | Add-Member -NotePropertyName EngineHash -NotePropertyValue (File-Sha256 (Join-Path $PSScriptRoot 'Core.ps1')) -Force
 Save-Cache $s;return $s
}
