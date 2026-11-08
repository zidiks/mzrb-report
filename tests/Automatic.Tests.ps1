$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'app\Core.ps1');. (Join-Path $root 'app\State.ps1')
$script:Count=0
function Assert($ok,$message){if(-not $ok){throw ('FAILED: '+$message)};$script:Count++}
function Matrix(){return @{Values=(New-Object 'object[,]' 30,36);FirstRow=1;FirstColumn=1;LastRow=29;LastColumn=35}}
function Put($m,$r,$pairs){for($i=0;$i -lt $pairs.Count;$i+=2){$m.Values[$r,[int]$pairs[$i]]=$pairs[$i+1]}}
$b=Matrix;Put $b 4 @(17,'за период с 01.10.2026 по 31.10.2026');Put $b 6 @(3,'Наименование лекарственного средства');Put $b 7 @(11,'номер')
Put $b 9 @(2,1,3,'Препарат A',7,'Компания',11,'DEL-2026',14,'01.01.2026',16,1000,28,50,33,3)
Put $b 10 @(2,2,3,'Препарат B',7,'Компания',11,'DEL-2026',14,'01.01.2026',16,500,28,25)
Put $b 12 @(14,'ВСЕГО:')
function Auto($raw){$s=Analyze-Matrix $raw $b 'source.xlsx' 'basis.xls' '01.11.26' ([datetime]'2026-10-01') $false $true $true;$s | Add-Member NoteProperty BasisPath 'basis.xls';Set-AutomaticBasisMatches $s (Read-Basis $b ([datetime]'2026-10-01'));return $s}
function Plan($s){return Merge-Basis $s (Read-Basis $b ([datetime]'2026-10-01'))}
$raw=Matrix;Put $raw 1 @(22,'Отпущено ЛПУ');Put $raw 5 @(1,1,7,'DEL-2026',8,'Компания',9,'Название из расчётов',22,'Октябрь',23,1500);Put $raw 6 @(7,'01.01.2026',22,'ком.возн.10',23,200);Put $raw 7 @(1,'Итого')
$s=Auto $raw;$plan=Plan $s
Assert ($plan.Rows -is [object[]] -and $plan.Changes -is [object[]] -and $plan.Notes -is [object[]] -and @($plan.Notes).Count -eq $plan.Notes.Count) 'Public plan collections can be consumed as arrays without binder errors'
Assert ($s.Lines.Count -eq 1) 'Do not invent source drugs from all basis names'
Assert ($plan.MatchedContracts -eq 1 -and $plan.PositionCount -eq 2 -and $plan.Rows.Count -eq 2) 'Contract matching preserves basis membership'
Assert ($plan.Rows[0].Values[5] -eq 1000 -and $plan.Rows[1].Values[5] -eq 500) 'Preserve matching per-drug shipment split'
Assert ($plan.Rows[0].Values[10] -eq 200 -and $plan.Rows[1].Values[10] -eq 0) 'Record contract fee once in first populated row'
Assert ($plan.Rows[0].Values[12] -eq 0) 'Calculated monthly zero replaces stale basis expense'
Assert (-not (Needs-Review $s) -and @(Validate-Session $s $true).Count -eq 0) 'Automatic export has no mandatory wizard and retains source conservation validation'
$raw.Values[6,23]=75;$s=Auto $raw;$plan=Plan $s
Assert ($plan.Rows[0].Values[10] -eq 50 -and $plan.Rows[1].Values[10] -eq 25) 'Keep a valid existing fee allocation, not only its total'
$raw.Values[6,9]='Второй препарат';$raw.Values[6,23]=200;$s=Auto $raw;$plan=Plan $s
Assert ($s.Lines.Count -eq 2 -and $plan.Rows[0].Values[10] -eq 200) 'Aggregate multiple source drug labels by contract without per-drug confirmations'
$raw.Values[5,8]='Компания';$raw.Values[6,8]='Компания';$s=Auto $raw
Assert ($s.Lines[0].Supplier -eq 'Компания') 'Deduplicate supplier fragments'
$raw.Values[5,7]='Дистриб.контр.';$b.Values[9,11]='ДИСТРИБЬЮТОРСКИЙ КОНТРАКТ';$b.Values[10,11]='ДИСТРИБЬЮТОРСКИЙ КОНТРАКТ'
$s=Auto $raw;$plan=Plan $s
Assert ($plan.MatchedContracts -eq 1 -and $plan.Rows[0].Values[10] -eq 200) 'Distributor contract abbreviations match full 1C contract names'
$raw.Values[5,7]='DIFFERENT-123';$s=Auto $raw;$plan=Plan $s
Assert ($plan.MatchedContracts -eq 0 -and $plan.Rows[0].Values[10] -eq 50 -and $plan.Rows[0].Values[12] -eq 3) 'A contract absent from the ledger is preserved'
Assert (-not (Needs-Review $s) -and $plan.Skipped.Count -gt 0) 'Unrelated source contracts never create mandatory tasks'
$raw.Values[5,7]='Дистриб.контр.';$raw.Values[5,23]=1600;$s=Auto $raw;$plan=Plan $s
Assert ($plan.Rows[0].Values[5] -eq 1000 -and $plan.Rows[1].Values[5] -eq 500 -and @($plan.Notes | Where-Object {$_.Relevant}).Count -eq 1) 'Unknown drug-level shipment allocation is reported, not guessed'
$raw.Values[5,23]=1500;$raw.Values[5,22]='Сентябрь';$raw.Values[6,22]='ком.возн.09';$s=Auto $raw;$plan=Plan $s
Assert ($s.Groups.Count -eq 0 -and $s.AllContracts.Count -eq 1 -and $plan.MatchedContracts -eq 1) 'Recognize contracts even with no monthly events'
Assert ($plan.Rows[0].Values[5] -eq 0 -and $plan.Rows[1].Values[5] -eq 0 -and $plan.Total -eq 0) 'A contract present in ledger with no monthly values has computed zeros'
# Same contract number with different dates must use the source date, not mix suppliers/years.
$raw.Values[5,7]='DEL-2026';$b.Values[9,11]='DEL-2026';$b.Values[10,11]='DEL-2026';$b.Values[10,14]='02.01.2026';$raw.Values[5,22]='Октябрь';$raw.Values[6,22]='ком.возн.10';$raw.Values[6,7]='01.01.2026'
$s=Auto $raw;$plan=Plan $s
Assert ($plan.Rows[0].Values[5] -eq 1500 -and $plan.Rows[1].Values[5] -eq 500) 'Use contract date to resolve a reused number'
$raw.Values[6,7]='';$s=Auto $raw;$plan=Plan $s
Assert ($plan.MatchedContracts -eq 0 -and $plan.Rows[0].Values[5] -eq 1000) 'Reused number without a unique date/supplier remains unchanged'
# Full signed redistribution folds into an existing row, adding no positions.
$b.Values[10,14]='01.01.2026';$raw.Values[6,7]='01.01.2026';$raw.Values[6,22]='Октябрь(перер.)';$raw.Values[6,23]=-25
$s=Auto $raw;$plan=Plan $s
Assert ($plan.Rows.Count -eq 2 -and $plan.Rows[0].Values[5] -eq 975 -and $plan.Rows[1].Values[5] -eq 500) 'Fold redistribution into a basis row without altering the remaining shipment split'
$s.Lines[0].Transfer=1
Assert (@(Validate-Session $s $true).Count -gt 0) 'Still reject a broken source total'
"$script:Count automatic contract checks passed"
