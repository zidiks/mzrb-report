$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'app\Core.ps1')
$count=0
function Assert($ok,$message){if(-not $ok){throw ('FAILED: '+$message)};$script:count++}
function Row($id,$name,$supplier,$contract,$amount){$v=New-Object 'object[]' 13;$v[0]=[string]$id;$v[1]=$name;$v[2]=$supplier;$v[3]=$contract;$v[4]='01.01.2026';$v[5]=[decimal]$amount;return [pscustomobject]@{Row=$id;Identity=('basis:'+$id);ParentKey='';ParentRow=0;IsMain=$true;IsRedistribution=$false;Values=$v}}
$rows=New-Object Collections.Generic.List[object]
$a=Row 9 'Янтарь' 'Company, РФ' 'C-1' 100;$a.Values[6]='Клиника';$a.Values[7]=5;$rows.Add($a)
$i=Row 10 '' '' '' 0;$i.IsMain=$false;$i.ParentKey=$a.Identity;$i.Values[0]=$null;$i.Values[6]='"';$i.Values[7]=10;$rows.Add($i)
$b=Row 11 'Beta' 'Поставщик, РБ' 'B-1' 20;$b.Values[6]='—';$rows.Add($b)
$c=Row 12 'Alpha' 'Company, РФ' 'C-2' 50;$rows.Add($c)
$d=Row 13 'Абрикос' 'Company, РФ' 'C-1' 30;$rows.Add($d)
$e=Row 14 'Alpha' 'Another, РФ' 'C-3' 60;$rows.Add($e)
$changes=New-Object Collections.Generic.List[object];Add-OutputRedistribution $rows @($a,$d) $a ([decimal]-5) $changes @(1)
$s=[pscustomobject]@{AllContracts=@()};$layout=Finalize-OutputRows $rows $s
$main=@($layout.Rows | Where-Object {$_.IsMain})
Assert (($main | ForEach-Object {$_.Values[3]}) -join ',' -eq 'C-3,C-2,C-1,C-1,B-1') 'Nonresidents first; Latin before Cyrillic; supplier breaks name ties'
Assert ($main[2].Values[1] -eq 'Абрикос' -and $main[3].Values[1] -eq 'Янтарь') 'Drug order inside one contract'
Assert ($layout.ContractBlocks.Count -eq 4) 'Separated occurrences of a contract form one block'
$group=@($layout.ContractBlocks | Where-Object {$_.Key.StartsWith('C-1|')})[0]
Assert ($layout.Rows[$group.EndIndex].IsRedistribution -and $group.DataEndIndex -lt $group.EndIndex) 'Redistribution is last after all contract positions and institutions'
Assert (($main | ForEach-Object {$_.Values[0]}) -join ',' -eq '1,2,3,4,5') 'Numbering occurs after final grouping and sorting'
Assert ($i.Values[6] -eq 'Клиника' -and $i.Values[7] -eq 10) 'Expand a funded ditto mark before sorting'
Assert ($b.Values[6] -eq $null) 'Remove empty institution dash'
$x=Row 15 'X' 'Company, РФ' 'X' 0;$x.Values[6]='Клиника — центр';Clean-InstitutionNames @($x)
Assert ($x.Values[6] -eq 'Клиника — центр') 'Keep punctuation inside real institution names'
Assert ((Supplier-Residency 'English Company, РБ') -eq 1 -and (Supplier-Residency 'Кириллическая компания, РФ') -eq 0) 'Residency is not inferred from alphabet'
Assert ((Name-ScriptRank '«Alpha»') -eq 0 -and (Name-ScriptRank '«Альфа»') -eq 1) 'Ignore leading quotes for alphabet priority'
$total=[decimal]0;foreach($r in $layout.Rows){foreach($j in @(5,8,9,10,11,12)){$total+=Money $r.Values[$j]}}
$plan=[pscustomobject]@{Rows=$layout.Rows;ContractBlocks=$layout.ContractBlocks;Total=$total}
$p=Get-ReportPresentation $plan
Assert ($total -eq 255 -and $p.RedistributionTotal -eq -5 -and $p.TotalBeforeRedistribution -eq 260) 'Footer reconciles net total and redistribution without counting it twice'
Assert ($p.Redistributions.Count -eq 1 -and $p.Redistributions[0].TableRow -eq (9+$group.EndIndex)) 'Footer references the real table redistribution row'
Assert ($p.FooterHeading -gt $p.GrandRow -and $p.BeforeRow -gt $p.FooterSum -and $p.EndRow -eq $p.BeforeRow) 'Footer is below main totals and included in print range'
Assert ($p.NumberFormat -eq '#,##0.00') 'Use a standard numeric format with grouping and exactly two decimals, without locale extensions'
Assert ($p.Merges.Count -eq 5) 'Merge safe contract metadata and per-drug labels'
foreach($m in $p.Merges){Assert ($m.Column -le 5 -and $m.LastRow -lt (9+$group.EndIndex)) 'Never merge financial data or the redistribution label'}
# A month without redistribution still has a valid zero summary and unchanged before-total.
$empty=[pscustomobject]@{Rows=@($e);ContractBlocks=@();Total=[decimal]60};$p0=Get-ReportPresentation $empty
Assert ($p0.Redistributions.Count -eq 0 -and $p0.RedistributionTotal -eq 0 -and $p0.TotalBeforeRedistribution -eq 60) 'Zero redistribution month is valid'
# Model Excel's rejection of a range that intersects only F of merged F:G.
function Test-FormatRange($sheet,[string]$address,[bool]$partial=$false){
 $range=[pscustomobject]@{Sheet=$sheet;Address=$address;Partial=$partial;StoredFormat=''}
 $range | Add-Member ScriptProperty NumberFormat {$this.StoredFormat} {
  param($value)
  if($this.Partial){throw 'Cannot set NumberFormat on part of a merged cell'}
  $this.StoredFormat=$value;$this.Sheet.Writes.Add($this.Address)
 }
 return $range
}
$fake=[pscustomobject]@{Writes=(New-Object Collections.Generic.List[string]);GrandRow=$p.GrandRow}
$fake | Add-Member ScriptMethod Range {
 param($address)
 $range=Test-FormatRange $this $address ($address -eq ('F9:F'+$this.GrandRow))
 if($address -eq ('F'+$this.GrandRow)){$range | Add-Member NoteProperty MergeArea (Test-FormatRange $this ('F'+$this.GrandRow+':G'+$this.GrandRow))}
 return $range
}
$rejected=$false;try{$fake.Range('F9:F'+$p.GrandRow).NumberFormat=$p.NumberFormat}catch{$rejected=$true}
Assert $rejected 'Regression fixture rejects partial access to the inherited merged grand total'
Set-ReportMoneyFormats $fake $p
Assert ($fake.Writes.Count -eq 8 -and $fake.Writes.Contains('F'+$p.GrandRow+':G'+$p.GrandRow)) 'Format the entire merged grand total separately'
foreach($column in @('F','H','I','J','K','L','M')){Assert ($fake.Writes.Contains($column+'9:'+$column+$p.TotalRow)) 'Column formatting stops at subtotal before the merged grand-total row'}
# Simulate COM rejecting invariant format on a localized Office installation.
function Local-FormatFixture([bool]$system,[string]$decimal,[string]$group,[bool]$rejectLocal=$false){
 $app=[pscustomobject]@{UseSystemSeparators=$system;DecimalSeparator=$decimal;ThousandsSeparator=$group}
 $app | Add-Member ScriptMethod International {param($index);if($index -eq 3){return ','};if($index -eq 4){return [string][char]160};throw 'Unexpected International index'}
 $r=[pscustomobject]@{Application=$app;Accepted='';RejectLocal=$rejectLocal;InvariantAttempts=0;LocalAttempts=0}
 $r | Add-Member ScriptProperty NumberFormat {''} {param($value);$this.InvariantAttempts++;throw 'NumberFormat COM rejection'}
 $r | Add-Member ScriptProperty NumberFormatLocal {$this.Accepted} {param($value);$this.LocalAttempts++;if($this.RejectLocal){throw 'NumberFormatLocal COM rejection'};$this.Accepted=$value}
 return $r
}
$r=Local-FormatFixture $true '.' ',';Set-ExcelMoneyFormat $r 'F9:F12'
Assert ($r.Accepted -eq ('#'+[char]160+'##0,00') -and $r.InvariantAttempts -eq 1 -and $r.LocalAttempts -eq 1) 'Localized fallback uses active system separators reported by Excel'
$r=Local-FormatFixture $false ',' ' ';Set-ExcelMoneyFormat $r 'H9:H12'
Assert ($r.Accepted -eq '# ##0,00') 'Localized fallback honors custom Excel separators'
$r=Local-FormatFixture $false '.' ',';Set-ExcelMoneyFormat $r 'I9:I12'
Assert ($r.Accepted -eq '#,##0.00') 'Localized fallback also supports English Excel'
$r=Local-FormatFixture $true ',' ' ' $true;$detail='';try{Set-ExcelMoneyFormat $r 'K9:K12'}catch{$detail=$_.Exception.Message}
Assert ($detail.Contains('K9:K12') -and $detail.Contains('NumberFormat COM rejection') -and $detail.Contains('NumberFormatLocal COM rejection')) 'If both formats fail, stop with range address and both errors instead of ignoring formatting failure'
"$count presentation checks passed"
