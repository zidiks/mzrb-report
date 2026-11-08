$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'app\Core.ps1')
$script:Count=0
function Assert($condition,[string]$message){if(-not $condition){throw "FAILED: $message"};$script:Count++}
function Matrix([int]$rows,[int]$cols){return @{Values=(New-Object 'object[,]' $rows,$cols);FirstRow=1;FirstColumn=1;LastRow=$rows-1;LastColumn=$cols-1}}
function Put($m,[int]$r,[object[]]$pairs){for($i=0;$i -lt $pairs.Length;$i+=2){$m.Values[$r,[int]$pairs[$i]]=$pairs[$i+1]}}
$template=Matrix 207 35
Put $template 6 @(3,'Наименование лекарственного средства');Put $template 7 @(11,'номер')
$raw=Matrix 25 35;Put $raw 1 @(22,'Отпущено ЛПУ')
Put $raw 5 @(1,'190.1',7,'A-1',8,'Фирма',9,'Препарат',14,46303.0,18,100.10,20,200.20,22,'ком.возн.10',23,100.10)
Put $raw 6 @(7,'01.02.2026',22,'расх.10',23,200.20)
Put $raw 7 @(22,'расх.10',23,3.40);Put $raw 8 @(22,'курс.разн.10',23,-5.60);Put $raw 9 @(22,'Октябрь',23,1000);Put $raw 10 @(22,'Октябрь(перер.)',23,-20)
Put $raw 11 @(14,46303.0,18,999999,20,888888,22,'Сентябрь',23,99999);Put $raw 12 @(1,'Итого',23,999999)
Put $raw 13 @(1,191,7,'A-2',8,'Поставщик',9,'Первый',22,'Октябрь',23,50);Put $raw 14 @(7,'02.02.2026',9,'Второй');Put $raw 15 @(1,'Итого',23,50)
Put $raw 16 @(1,192,7,'A-3',9,'Неактивный',22,'Сентябрь',23,999);Put $raw 17 @(1,'Итого',23,999)
$s=Analyze-Matrix $raw $template 'source' 'template' '01.11.26' ([datetime]'2026-10-01') $false
Assert ($s.Groups.Count -eq 2 -and $s.Lines.Count -eq 3) 'Decimal contract IDs, inactive months, separate drug rows'
Assert ($s.Lines[0].Conversion -eq 100.10 -and $s.Lines[0].Commission -eq 200.20 -and $s.Lines[0].Other -eq 3.40) 'Swapped expense labels use numeric evidence'
Assert ($s.Lines[0].Fx -eq -5.60 -and $s.Lines[0].Transfer -eq 1000 -and $s.Lines[0].Redistribution -eq -20) 'Signed values, shipment and redistribution'
Assert ((Source-Total $s) -eq (Session-Total $s)) 'No subtotal duplication'
Assert ($s.Lines[2].Transfer -eq 0) 'Multi-drug allocation not duplicated'
$baseline=($s.Lines | ForEach-Object {foreach($f in $script:Fields){Fixed $_.$f}}) -join '|'
foreach($otherDate in @(45000.0,47000.0,$null)){
 for($r=2;$r -lt 25;$r++){foreach($c in @(4,12,14)){$raw.Values[$r,$c]=$otherDate}}
 $changed=Analyze-Matrix $raw $template 'source' 'template' '01.11.26' ([datetime]'2026-10-01') $false
 Assert ((($changed.Lines | ForEach-Object {foreach($f in $script:Fields){Fixed $_.$f}}) -join '|') -eq $baseline) 'Other dates never change monthly amounts'
}
Assert (@(Validate-Session $s $false).Count -eq 0) 'Valid draft allowed'
Assert (@(Validate-Session $s $true).Count -gt 0) 'Unreviewed final rejected'
$s.Lines[1].Transfer=[decimal]30;$s.Lines[2].Transfer=[decimal]20;foreach($l in $s.Lines){$l.Reviewed=$true}
Assert (@(Validate-Session $s $true).Count -eq 0) 'Conserved multi-drug split accepted'
$s.Lines[2].Transfer=[decimal]50
Assert (@(Validate-Session $s $false).Count -gt 0) 'Duplicated allocation rejected'
Assert ((Sheet-Period '01.01.27').ToString('yyyy-MM') -eq '2026-12') 'January selects December of previous year'
Assert ((Sheet-Period '01.11.2026 (МЗ)').ToString('yyyy-MM') -eq '2026-10') 'Four-digit year'
Assert ((Label-Type 'Cентябрь').Month -eq 9 -and (Label-Type 'Септябрь').Month -eq 9) 'Month spelling variants'
Assert ((Label-Type 'тов.скидка08').Category -eq 'Discount' -and (Label-Type 'НДС10').Category -eq 'Vat') 'Expense label variants'
Assert ((Label-Type 'Тов.скидка').Month -eq 0) 'No invented month for unlabelled discount'
$commission=Matrix 10 35;Put $commission 1 @(22,'Отпущено ЛПУ');Put $commission 5 @(1,1,7,'A-5',9,'Препарат',18,500,20,50,22,'ком.возн.10',23,50);Put $commission 6 @(7,'01.01.2026');Put $commission 7 @(1,'Итого')
$cs=Analyze-Matrix $commission $template 'source' 'template' '01.11.26' ([datetime]'2026-10-01') $false
Assert ($cs.Lines[0].Commission -eq 50 -and $cs.Lines[0].Conversion -eq 0 -and $cs.Lines[0].Other -eq 0) 'Commission alone must not include unlabelled currency costs'
$publicOctober=Analyze-Matrix $raw $template 'source' 'template' '01.11.26' ([datetime]'2026-10-01') $true
Assert ($publicOctober.Lines.Count -gt 0 -and $null -eq $publicOctober.Reference) 'Public install handles October without a private profile'
"$script:Count checks passed"
