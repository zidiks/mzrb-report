$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'app\Core.ps1')
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'app\State.ps1')
$script:Checks=0
function Assert($ok,[string]$message){if(-not $ok){throw ('FAILED: '+$message)};$script:Checks++}
function Reject([scriptblock]$action,[string]$message){$failed=$false;try{& $action | Out-Null}catch{$failed=$true};Assert $failed $message}
function Matrix([int]$n=20){return @{Values=(New-Object 'object[,]' $n,37);FirstRow=1;FirstColumn=1;LastRow=$n-1;LastColumn=36}}
function Put($m,[int]$r,[object[]]$values){for($i=0;$i -lt $values.Count;$i+=2){$m.Values[$r,[int]$values[$i]]=$values[$i+1]}}
$base=Matrix
Put $base 4 @(17,'за период с 01.10.2026 по 31.10.2026')
Put $base 6 @(3,'Наименование лекарственного средства');Put $base 7 @(11,'номер')
Put $base 9 @(2,'1',3,'Препарат A',7,'Поставщик',11,'A-1',14,'01.01.2026',16,'1 000,00',18,'Учреждение 1',21,'250,00',28,'50,00',33,'3,00')
Put $base 10 @(18,'Учреждение 2',21,'100,00')
Put $base 11 @(2,'2',3,'Препарат B',7,'Поставщик',11,'A-2',14,'02.01.2026',16,200,25,-2)
Put $base 12 @(16,1200,21,350,28,50,33,3);Put $base 13 @(14,'ВСЕГО:',16,1251)
$raw=Matrix
Put $raw 1 @(22,'Отпущено ЛПУ');Put $raw 5 @(1,1,7,'A-1',8,'Поставщик',9,'Препарат A',22,'ком.возн.10',23,0);Put $raw 6 @(7,'01.01.2026',22,'Октябрь',23,900);Put $raw 7 @(1,'Итого')
$period=[datetime]'2026-10-01'
$basis=Read-Basis $base $period
function Session($matrix){$s=Analyze-Matrix $matrix $base 'ledger.xlsx' 'basis.xls' '01.11.26' $period $false $true;$s | Add-Member NoteProperty BasisPath 'basis.xls';Set-BasisMatches $s $basis;return $s}
$s=Session $raw;$plan=Merge-Basis $s $basis
Assert ($s.Lines[0].BasisRow -eq 9) 'Exact contract, name, supplier and date matched'
Assert ($plan.Rows.Count -eq 3) 'All physical base rows including continuation kept'
Assert ($plan.Rows[0].Values[5] -eq 900) 'Shipment replaced, not added'
Assert ($plan.Rows[0].Values[10] -eq 0) 'Explicit zero replaces commission'
Assert ($plan.Rows[0].Values[12] -eq 3) 'Missing other category preserves base amount'
Assert ($plan.Rows[0].Values[6] -eq 'Учреждение 1' -and $plan.Rows[1].Values[6] -eq 'Учреждение 2') 'Institution names and continuation preserved'
Assert ($plan.Rows[0].Values[7] -eq 250 -and $plan.Rows[1].Values[7] -eq 100) 'Institution amounts preserved'
Assert ($plan.Rows[2].Values[5] -eq 200 -and $plan.Rows[2].Values[9] -eq -2) 'Unmatched base drug and signed amount preserved'
Assert ($plan.Total -eq 1101) 'Grand total excludes institution subtotal'
Assert ((Cell $base 9 28) -eq '50,00' -and (Cell $base 9 16) -eq '1 000,00') 'Original base matrix untouched'
$raw.Values[5,23]=$null;$s=Session $raw;$plan=Merge-Basis $s $basis
Assert ($plan.Rows[0].Values[10] -eq 50) 'Blank amount does not become zero override'
$raw.Values[6,23]=$null;$s=Session $raw;$plan=Merge-Basis $s $basis
Assert ($s.Lines.Count -eq 0 -and $plan.Total -eq 1251) 'No monthly values preserves complete base'
$raw.Values[6,23]=900;$raw.Values[5,23]=0
Put $raw 9 @(1,2,7,'A-3',8,'Другой поставщик',9,'Новый препарат',22,'Октябрь',23,30);Put $raw 10 @(7,'03.01.2026');Put $raw 11 @(1,'Итого')
$s=Session $raw
Assert ($s.Lines[-1].BasisRow -eq -1 -and -not $s.Lines[-1].Reviewed) 'Contract absent from basis requires explicit review'
$s.Lines[-1].BasisRow=0;$plan=Merge-Basis $s $basis
Assert ($plan.PositionCount -eq 2 -and $plan.Rows.Count -eq 3 -and $plan.Total -eq 1101 -and $plan.Skipped.Count -eq 1) 'Explicit skip preserves exact base position list'
Put $raw 12 @(1,3,7,'A-1',8,'Поставщик',9,'Препарат A',22,'Октябрь',23,10);Put $raw 13 @(7,'01.01.2026');Put $raw 14 @(1,'Итого')
$s=Session $raw;$s.Lines[1].BasisRow=0;$plan=Merge-Basis $s $basis
Assert ($plan.Rows[0].Values[5] -eq 910) 'Multiple ledger blocks aggregate into one replacement'
Put $raw 15 @(1,4,7,'A-1',8,'Поставщик',9,'Другое название',22,'Октябрь',23,10);Put $raw 16 @(7,'01.01.2026');Put $raw 17 @(1,'Итого')
$s=Session $raw;$s.Lines[1].BasisRow=0
Assert ($s.Lines[-1].BasisRow -eq -1 -and -not $s.Lines[-1].Reviewed) 'Same contract with different drug requires mapping review'
Reject {Merge-Basis $s $basis} 'Unresolved mapping blocks export'
$s.Lines[-1].BasisRow=0;$plan=Merge-Basis $s $basis
Assert ($plan.PositionCount -eq 2 -and $plan.Skipped.Count -eq 2) 'No new positions may come from ledger'
$s.Lines[-1].BasisRow=11
Reject {Merge-Basis $s $basis} 'Cannot overwrite a different contract'
Reject {Read-Basis $base ([datetime]'2026-09-01')} 'Mismatched base month rejected'
$base.Values[9,35]='Unexpected data'
Reject {Read-Basis $base $period} 'Nonempty phantom column is not silently discarded'
$base.Values[9,35]=$null
# Redistribution is folded into shipment; exact basis rows are retained.
$redis=Matrix;Put $redis 1 @(22,'Отпущено ЛПУ');Put $redis 5 @(1,1,7,'A-1',8,'Поставщик',9,'Препарат A',22,'Октябрь(перер.)',23,-25);Put $redis 6 @(7,'01.01.2026');Put $redis 7 @(1,'Итого')
$s=Session $redis;$plan=Merge-Basis $s $basis
Assert ($plan.Rows.Count -eq 3 -and $plan.Rows[0].Values[5] -eq 975) 'Redistribution folded into existing shipment row'
Assert ($plan.Rows[1].Values[7] -eq 100 -and $plan.Total -eq 1226) 'Redistribution preserves institution data and contributes once'
# Shipment replacement is applied before signed redistribution, independent of map order.
$redis.Values[6,22]='Октябрь';$redis.Values[6,23]=900
$s=Session $redis;$plan=Merge-Basis $s $basis
Assert ($plan.Rows.Count -eq 3 -and $plan.Rows[0].Values[5] -eq 875) 'Shipment replacement and redistribution fold exactly once'
# A duplicate target must be reviewed rather than selecting the first matching contract/drug.
$duplicate=[pscustomobject]@{Rows=$basis.Rows;Main=@($basis.Main[0],$basis.Main[0],$basis.Main[1]);TotalRow=$basis.TotalRow;GrandRow=$basis.GrandRow}
$s=Session $redis;Set-BasisMatches $s $duplicate
Assert ($s.Lines[0].BasisRow -eq -1 -and -not $s.Lines[0].Reviewed) 'Duplicate basis match requires review'
$s.Lines[0].BasisRow=9
# Preserve manual row choices and invalidate cache independently for each basis book/sheet.
Assert ((Cache-Path 'a.xlsx' 'sheet' '2026-10' 'b.xls' 'one') -ne (Cache-Path 'a.xlsx' 'sheet' '2026-10' 'b.xls' 'two')) 'Basis sheet belongs to cache identity'
Assert ((Cache-Path 'a.xlsx' 'sheet' '2026-10' 'b.xls' 'one') -ne (Cache-Path 'a.xlsx' 'sheet' '2026-10' 'c.xls' 'one')) 'Basis file belongs to cache identity'
$temp=Join-Path ([IO.Path]::GetTempPath()) ('merge-test-'+[guid]::NewGuid().ToString('N')+'.json')
try{Save-Session $s $temp;$restored=Restore-Session $temp;Assert ($restored.BasisPath -eq 'basis.xls' -and $restored.Lines[0].BasisRow -eq 9) 'Two-input review round trip'}finally{if(Test-Path -LiteralPath $temp){Remove-Item -LiteralPath $temp}}
"$script:Checks merge checks passed"
