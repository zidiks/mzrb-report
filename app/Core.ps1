$ErrorActionPreference='Stop'
$script:Invariant=[Globalization.CultureInfo]::InvariantCulture
$script:Fields=@('Transfer','Redistribution','Conversion','Fx','Commission','Discount','Other')
$script:Months=@('январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь')
function Text($v) { if($null -eq $v){return ''};return [Convert]::ToString($v,$script:Invariant).Trim() }
function Money($v) {
 if($null -eq $v -or (Text $v) -eq ''){return [decimal]0}
 $str=(Text $v).Replace([string][char]160,'').Replace([string][char]8239,'').Replace(' ','').Replace(',','.')
 return [decimal]::Round([decimal]::Parse($str,[Globalization.NumberStyles]::Float,$script:Invariant),2,[MidpointRounding]::AwayFromZero)
}
function Fixed($v){return (Money $v).ToString('F2',$script:Invariant)}
function Normal($v){return [regex]::Replace((Text $v).ToLowerInvariant().Replace('ё','е'),'[^\p{L}\p{N}]','')}
function Parse-Date($v) {
 if($v -is [datetime]){return $v}
 if($v -is [double] -or $v -is [int]){if($v -gt 20000 -and $v -lt 100000){return [datetime]::FromOADate($v)};return $null}
 $m=[regex]::Match((Text $v),'(?<!\d)(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4}|\d{2})(?!\d)')
 if($m.Success){$year=[int]$m.Groups[3].Value;if($year -lt 100){$year+=2000};try{return New-Object datetime($year,[int]$m.Groups[2].Value,[int]$m.Groups[1].Value)}catch{}}
 return $null
}
function Sheet-Period([string]$name){$date=Parse-Date $name;if($null -eq $date){throw 'В названии вкладки нет даты. Укажите месяц вручную.'};return (New-Object datetime($date.Year,$date.Month,1)).AddMonths(-1)}
function Label-Type($value){
 $date=Parse-Date $value
 if($null -ne $date){return @{Month=$date.Month;Year=$date.Year;Category='Transfer'}}
 $s=(Text $value).ToLowerInvariant().Replace('c','с').Replace('o','о').Replace('a','а').Replace('e','е').Replace('септябрь','сентябрь')
 for($i=0;$i -lt 12;$i++){if($s.StartsWith($script:Months[$i])){return @{Month=$i+1;Year=0;Category=$(if($s.Contains('перер')){'Redistribution'}else{'Transfer'})}}}
 $category='Unknown'
 if($s.StartsWith('курс')){$category='Fx'}elseif($s.StartsWith('ком')){$category='Commission'}elseif($s.StartsWith('расх')){$category='Expenses'}elseif($s.StartsWith('окр')){$category='Other'}elseif($s.StartsWith('ндс')){$category='Vat'}elseif($s.Contains('скид') -or $s.Contains('кр.') -or $s.Contains('нота')){$category='Discount'}
 $m=[regex]::Match($s,'(?<!\d)(0?[1-9]|1[0-2])(?!\d)')
 return @{Month=$(if($m.Success){[int]$m.Value}else{0});Year=0;Category=$category}
}
function New-Excel {
 $e=New-Object -ComObject Excel.Application
 $e.Visible=$false;$e.DisplayAlerts=$false;$e.AutomationSecurity=3;$e.AskToUpdateLinks=$false
 return $e
}
# Only the bundled layout (or its output copy) may be repaired. Never repair a user's source.
function Open-Layout($excel,[string]$path,[bool]$readOnly=$true){
 try{return $excel.Workbooks.Open($path,0,$readOnly)}catch{
  $normalError=$_.Exception.Message
  $missing=[Type]::Missing
  try{
   # CorruptLoad=xlRepairFile (1), supported by Excel 2010. Keep Office validation enabled.
   return $excel.Workbooks.Open($path,0,$readOnly,$missing,$missing,$missing,$true,$missing,$missing,$false,$false,$missing,$false,$true,1)
  }catch{throw ('Excel не смог открыть встроенный шаблон отчёта «'+[IO.Path]::GetFileName($path)+'» (Excel '+$excel.Version+'). Исходный файл не изменён. Обычное открытие: '+$normalError+'. Восстановление: '+$_.Exception.Message)}
 }
}
function Close-Excel($excel){if($null -ne $excel){try{$excel.Quit()}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel);[GC]::Collect();[GC]::WaitForPendingFinalizers()}}}
function Read-Matrix($sheet){$used=$sheet.UsedRange;return @{Values=$used.Value2;FirstRow=[int]$used.Row;FirstColumn=[int]$used.Column;LastRow=[int]$used.Row+[int]$used.Rows.Count-1;LastColumn=[int]$used.Column+[int]$used.Columns.Count-1}}
function Cell($matrix,[int]$row,[int]$col){if($row -lt $matrix.FirstRow -or $row -gt $matrix.LastRow -or $col -lt $matrix.FirstColumn -or $col -gt $matrix.LastColumn){return $null};return $matrix.Values[($row-$matrix.FirstRow+1),($col-$matrix.FirstColumn+1)]}
function Read-SheetNames([string]$source){
 $e=$null;$book=$null
 try{$e=New-Excel;$book=$e.Workbooks.Open($source,0,$true);$names=@();for($i=1;$i -le $book.Worksheets.Count;$i++){$names+=[string]$book.Worksheets.Item($i).Name};return $names}
 finally{if($null -ne $book){$book.Close($false)};Close-Excel $e}
}
function Sum-Events($events,[string]$category){$sum=[decimal]0;foreach($ev in $events){if($ev.Category -eq $category){$sum+=$ev.Amount}};return $sum}
function Line-Total($line){$sum=[decimal]0;foreach($field in $script:Fields){$sum+=Money $line.$field};return $sum}
function Session-Total($s){$sum=[decimal]0;foreach($l in $s.Lines){$sum+=Line-Total $l};return $sum}
function Source-Total($s){$sum=[decimal]0;foreach($ev in $s.Events){$sum+=$ev.Amount};return $sum}
function Get-Hash([string]$value){$sha=[Security.Cryptography.SHA256]::Create();try{return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}}
function File-Sha256([string]$path){
 $stream=$null;$sha=$null
 try{
  $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
  $sha=[Security.Cryptography.SHA256]::Create()
  return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','')
 }finally{if($null -ne $stream){$stream.Dispose()};if($null -ne $sha){$sha.Dispose()}}
}
function Ledger-Hash($s){
 $b=New-Object Text.StringBuilder
 foreach($ev in $s.Events){[void]$b.Append("$($ev.BlockId)|$($ev.Row)|$($ev.Contract)|$($ev.Label)|$(Fixed $ev.Amount)`n")}
 foreach($ex in $s.Expenses){[void]$b.Append("E|$($ex.BlockId)|$($ex.Row)|$(Fixed $ex.Total17)|$(Fixed $ex.Conversion18)|$(Fixed $ex.Vat19)|$(Fixed $ex.Commission20)|$(Fixed $ex.Other21)`n")}
 return Get-Hash $b.ToString()
}
function Template-Hash($matrix,$profile){
 $b=New-Object Text.StringBuilder
 for($r=$matrix.FirstRow;$r -le $matrix.LastRow;$r++){for($c=1;$c -le 34;$c++){
  if($r -ge 9 -and $r -le $profile.TotalRow -and $c -ge 18 -and $c -le 22){continue}
  $v=Cell $matrix $r $c;$text=Text $v;if($text -eq ''){continue}
  $prefix='s:'+$text;if($v -is [double] -or $v -is [int] -or $v -is [decimal] -or $v -is [long]){$prefix='n:'+(Fixed $v)}
  [void]$b.Append("$r|$c|$prefix`n")
 }}
 return Get-Hash $b.ToString()
}
function New-Line($group,[string]$name){
 return [pscustomobject]@{Group=$group.Id;SourceRow=$group.Row;Number=$group.Number;Name=$name;Supplier=$group.Supplier;Contract=$group.Contract;ContractDate=$group.ContractDate;Transfer=[decimal]0;Redistribution=[decimal]0;Conversion=[decimal]0;Fx=[decimal]0;Commission=[decimal]0;Discount=[decimal]0;Other=[decimal]0;Reviewed=$false;Note='';ReferenceRow=0;RedistributionRow=0}
}
function Analyze-Matrix($raw,$template,[string]$source,[string]$templatePath,[string]$sheet,[datetime]$period,[bool]$useReference=$true,[bool]$allowEmpty=$false,[bool]$basisDriven=$false){
 if(-not (Text (Cell $raw 1 22)).Contains('Отпущено')){
  $found=$false;for($r=1;$r -le [Math]::Min(7,$raw.LastRow);$r++){if((Text (Cell $raw $r 22)).Contains('Отпущено')){$found=$true}}
  if(-not $found){throw 'На выбранной вкладке не найден заголовок «Отпущено ЛПУ» в колонке 22.'}
 }
 if(-not (Text (Cell $template 6 3)).Contains('лекарственного') -or (Text (Cell $template 7 11)) -ne 'номер'){throw 'Повреждён встроенный формат отчёта. Восстановите папку программы.'}
 $refs=@{}
 for($r=9;$r -lt 203;$r++){$key=Normal (Cell $template $r 11);$name=Text (Cell $template $r 3);if($key -and $name){if(-not $refs.ContainsKey($key)){$refs[$key]=New-Object Collections.Generic.List[string]};if(-not $refs[$key].Contains($name)){$refs[$key].Add($name)}}}
 $namesPath=Join-Path $PSScriptRoot 'Data\Names.json'
 if(Test-Path -LiteralPath $namesPath){foreach($entry in (Get-Content -LiteralPath $namesPath -Raw -Encoding UTF8 | ConvertFrom-Json)){$refs[$entry.Key]=New-Object Collections.Generic.List[string];foreach($name in $entry.Names){$refs[$entry.Key].Add([string]$name)}}}
 $s=[pscustomobject]@{Version=4;Source=$source;Template=$templatePath;Sheet=$sheet;PeriodKey=$period.ToString('yyyy-MM');RequestReference=$useReference;Lines=(New-Object Collections.Generic.List[object]);Groups=(New-Object Collections.Generic.List[object]);Events=(New-Object Collections.Generic.List[object]);Expenses=(New-Object Collections.Generic.List[object]);Warnings=(New-Object Collections.Generic.List[string]);Reference=$null;Confirmed=$false;LedgerFingerprint='';TemplateFingerprint=''}
 $blocks=New-Object Collections.Generic.List[object];$block=$null
 for($r=$raw.FirstRow;$r -le $raw.LastRow;$r++){
  $no=Text (Cell $raw $r 1)
  if($no -match '^\d+([.,]\d+)?$' -and (Text (Cell $raw $r 7)) -ne ''){$block=[pscustomobject]@{Id=[string]$r;Row=$r;Number=$no;Contract=(Text (Cell $raw $r 7));ContractDate='';Supplier='';Rows=(New-Object Collections.Generic.List[int])};$blocks.Add($block)}
  if($null -eq $block){continue};if($no -eq 'Итого'){$block=$null;continue};$block.Rows.Add($r)
 }
 $allContracts=New-Object Collections.Generic.List[object]
 foreach($b in $blocks){
  $sup=@($b.Rows | ForEach-Object {Text (Cell $raw $_ 8)} | Where-Object {$_} | Select-Object -Unique)
  $allContracts.Add([pscustomobject]@{Id=$b.Id;Row=$b.Row;Contract=$b.Contract;ContractDate=(Text (Cell $raw ($b.Row+1) 7));Supplier=($sup -join ' ').Replace('- ','')})
 }
 $s | Add-Member NoteProperty AllContracts $allContracts.ToArray() -Force
 foreach($block in $blocks){
  $events=New-Object Collections.Generic.List[object]
  foreach($r in $block.Rows){
   $label=Text (Cell $raw $r 22);$amount=Money (Cell $raw $r 23)
   if($allowEmpty -and (Text (Cell $raw $r 23)) -eq ''){continue}
   if($label -eq ''){if($amount -ne 0){$s.Warnings.Add("Строка $r, $($block.Contract): сумма $amount без периода; не включена.")};continue}
   $type=Label-Type (Cell $raw $r 22)
   if($type.Month -eq 0){$s.Warnings.Add("Строка $r, $($block.Contract): месяц метки «$label» не определён, сумма $amount не включена.");continue}
   if($type.Month -ne $period.Month -or ($type.Year -ne 0 -and $type.Year -ne $period.Year)){continue}
   $event=[pscustomobject]@{BlockId=$block.Id;Row=$r;Contract=$block.Contract;Label=$label;Amount=$amount;Category=$type.Category};$events.Add($event);$s.Events.Add($event)
  }
  if($events.Count -eq 0){continue}
  $supplier=New-Object Collections.Generic.List[string];$names=New-Object Collections.Generic.List[string];$costs=New-Object Collections.Generic.List[object]
  foreach($r in $block.Rows){
   $sup=Text (Cell $raw $r 8);if($sup -and -not $supplier.Contains($sup)){$supplier.Add($sup)}
   $name=Text (Cell $raw $r 9);if($name -and $name.ToLowerInvariant().IndexOf('скидка') -lt 0 -and -not $names.Contains($name)){$names.Add($name)}
   $paymentDate=Parse-Date (Cell $raw $r 14)
   $ex=[pscustomobject]@{BlockId=$block.Id;Contract=$block.Contract;Row=$r;PaymentDate=$(if($null -ne $paymentDate){$paymentDate.ToString('dd.MM.yyyy')}else{Text (Cell $raw $r 14)});Total17=(Money (Cell $raw $r 17));Conversion18=(Money (Cell $raw $r 18));Vat19=(Money (Cell $raw $r 19));Commission20=(Money (Cell $raw $r 20));Other21=(Money (Cell $raw $r 21))}
   if($ex.Total17 -ne 0 -or $ex.Conversion18 -ne 0 -or $ex.Vat19 -ne 0 -or $ex.Commission20 -ne 0 -or $ex.Other21 -ne 0){$costs.Add($ex);$s.Expenses.Add($ex)}
  }
  $block.Supplier=($supplier -join ' ').Replace('- ','');$date=Parse-Date (Cell $raw ($block.Row+1) 7);$block.ContractDate=if($null -ne $date){$date.ToString('dd.MM.yyyy')}else{Text (Cell $raw ($block.Row+1) 7)}
  $notes=New-Object Collections.Generic.List[string];$review=$false
  $conversion=[decimal]0;$commission=Sum-Events $events 'Commission';$commEvents=@($events | Where-Object {$_.Category -eq 'Commission' -and $_.Amount -ne 0});$expEvents=@($events | Where-Object {$_.Category -eq 'Expenses' -and $_.Amount -ne 0})
  foreach($ev in $commEvents){
   $normal=@($costs | Where-Object {$_.Commission20 -eq $ev.Amount})
   $swapped=@($costs | Where-Object {$cost=$_;$cost.Conversion18 -eq $ev.Amount -and $cost.Commission20 -ne 0 -and @($expEvents | Where-Object {$_.Amount -eq $cost.Commission20}).Count -gt 0})
   if($normal.Count -gt 0){
    if($expEvents.Count -eq 0){continue}
    $components=@($normal | ForEach-Object {$_.Conversion18} | Select-Object -Unique)
    if($components.Count -eq 1){$conversion+=$components[0];$notes.Add("Комиссия из строки $($ev.Row) совпала с колонкой 20; покупка валюты $($components[0]) — из колонки 18.")}else{$review=$true;$notes.Add('Одинаковой комиссии соответствуют разные расходы покупки валюты. Проверьте распределение.')}
   }elseif($swapped.Count -gt 0){
    $pairs=@($swapped | Group-Object Conversion18,Commission20)
    if($pairs.Count -eq 1){$cost=$swapped[0];$conversion+=$cost.Conversion18;$commission+=$cost.Commission20-$ev.Amount;$notes.Add('Перепутанные метки подтверждены выбранными суммами и колонками 18/20.');$review=$true}else{$review=$true;$notes.Add('Несколько вариантов расшифровки колонок 18/20.')}
   }elseif($expEvents.Count -gt 0 -and @($costs | Where-Object {$_.Conversion18 -ne 0}).Count -gt 0){$review=$true;$notes.Add('Не найдена однозначная связь месячной комиссии с колонками 18/20.')}
  }
  if($commEvents.Count -eq 0){foreach($ev in $expEvents){
   $candidates=@($costs | Where-Object {$_.Conversion18 -ne 0 -and [Math]::Abs($_.Conversion18-$ev.Amount) -le 0.01} | ForEach-Object {$_.Conversion18} | Select-Object -Unique)
   $otherMatch=@($costs | Where-Object {$_.Other21 -ne 0 -and [Math]::Abs($_.Other21-$ev.Amount) -le 0.01}).Count -gt 0
   if($candidates.Count -eq 1 -and -not $otherMatch){$conversion+=$candidates[0]}elseif(@($costs | Where-Object {$_.Conversion18 -ne 0}).Count -gt 0){$review=$true;$notes.Add("Строка $($ev.Row): распределение расходов требует проверки по колонкам 18–21.")}
  }}
  $unsafeCosts=$review
  $known=@();$key=Normal $block.Contract;if($refs.ContainsKey($key)){$known=@($refs[$key].ToArray())}
  if($names.Count -gt 1){$joined=($names -join ' ').Replace('- ','-');$match=@($known | Where-Object {(Normal $_) -eq (Normal $joined)});$continuation=$true;for($i=1;$i -lt $names.Count;$i++){if($names[$i].Length -gt 2 -and -not [char]::IsLower($names[$i][0])){$continuation=$false}}
   if($match.Count -gt 0 -or $continuation -or @($names | Select-Object -SkipLast 1 | Where-Object {$_.EndsWith('-')}).Count -gt 0){$names.Clear();$names.Add($(if($match.Count){$match[0]}else{$joined}))}
  }
  if(-not $basisDriven -and $names.Count -eq 1 -and $known.Count -gt 1){$names.Clear();foreach($n in $known){$names.Add($n)};$review=$true}
  if($names.Count -eq 0){if(-not $basisDriven -and $known.Count){foreach($n in $known){$names.Add($n)}}else{$names.Add('Укажите наименование')};$review=$true}
  if($names.Count -gt 1){$review=$true;$notes.Add('Несколько препаратов: общая сумма временно стоит в первой строке. Распределите её, не увеличивая сумму договора.')}
  if($null -eq (Parse-Date $block.ContractDate)){$review=$true;$notes.Add('Проверьте дату договора в строке под его номером.')}
  $vat=Sum-Events $events 'Vat';if($vat -ne 0){$unsafeCosts=$true;$review=$true;$notes.Add('НДС включён в прочие расходы; требуется проверка.')}
  if(@($events | Where-Object {$_.Category -eq 'Unknown'}).Count){$unsafeCosts=$true;$review=$true;$notes.Add('Есть нераспознанная месячная метка: распределите сумму вручную.')}
  $expected=@((Sum-Events $events 'Transfer'),(Sum-Events $events 'Redistribution'),$conversion,(Sum-Events $events 'Fx'),$commission,(Sum-Events $events 'Discount'),((Sum-Events $events 'Commission')+(Sum-Events $events 'Expenses')+$vat+(Sum-Events $events 'Other')-$conversion-$commission))
  $g=[pscustomobject]@{Id=$block.Id;Row=$block.Row;Contract=$block.Contract;Expected=$expected;UnsafeCosts=$unsafeCosts;NeedsReview=$review;Note=($notes -join ' ')};$s.Groups.Add($g)
  for($i=0;$i -lt $names.Count;$i++){$line=New-Line $block $names[$i];$line.Note=$g.Note;$line.Reviewed=-not $review;if($i -eq 0){for($j=0;$j -lt 7;$j++){$line.($script:Fields[$j])=[decimal]$expected[$j]}};$s.Lines.Add($line)}
 }
 if($s.Lines.Count -eq 0 -and -not $allowEmpty){throw 'За выбранный месяц в колонке 22 не найдены записи. Проверьте вкладку и месяц.'}
 $s.LedgerFingerprint=Ledger-Hash $s
 if($useReference -and $s.PeriodKey -eq '2026-10' -and (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'Data\October.json'))){
  $p=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Data\October.json') -Raw -Encoding UTF8 | ConvertFrom-Json
  $s.TemplateFingerprint=Template-Hash $template $p
  $sameSourceFile=$p.SourceFileHash -and (Test-Path -LiteralPath $source -PathType Leaf) -and (File-Sha256 $source) -eq $p.SourceFileHash
  $layoutFingerprint=if($p.LayoutFingerprint){$p.LayoutFingerprint}else{$p.ReferenceFingerprint}
  if($sameSourceFile -and $s.LedgerFingerprint -eq $p.LedgerFingerprint -and $s.TemplateFingerprint -eq $layoutFingerprint){
   foreach($g in $s.Groups){$pg=@($p.Groups | Where-Object {$_.Id -eq $g.Id});if($pg.Count -eq 0){if(@($g.Expected | Where-Object {$_ -ne 0}).Count){throw 'В октябрьском профиле отсутствует активный договор.'};continue};for($i=0;$i -lt 7;$i++){if((Money $g.Expected[$i]) -ne (Money $pg[0].SourceAmounts[$i])){throw "Изменилась логика расчёта договора $($g.Contract); октябрьская сверка остановлена."}}}
   $s.Reference=$p;$s.Lines.Clear();$s.Groups.Clear()
   foreach($g in $p.Groups){$s.Groups.Add([pscustomobject]@{Id=$g.Id;Row=$g.SourceRow;Contract=$g.Contract;Expected=$g.ExpectedAmounts;NeedsReview=$false;Note=('Сверено с октябрьским образцом. '+$g.Reason+$(if($g.SplitFromReference){' Разбивка препаратов — из образца.'}else{''}))})}
   foreach($rec in $p.Records){$line=[pscustomobject]@{Group=$rec.Group;SourceRow=$rec.SourceRow;Number=$rec.Number;Name=$rec.Name;Supplier=$rec.Supplier;Contract=$rec.Contract;ContractDate=$rec.ContractDate;Transfer=(Money $rec.Amounts[0]);Redistribution=(Money $rec.Amounts[1]);Conversion=(Money $rec.Amounts[2]);Fx=(Money $rec.Amounts[3]);Commission=(Money $rec.Amounts[4]);Discount=(Money $rec.Amounts[5]);Other=(Money $rec.Amounts[6]);Reviewed=$true;Note=(@($s.Groups | Where-Object {$_.Id -eq $rec.Group})[0].Note);ReferenceRow=$rec.TargetRow;RedistributionRow=$rec.RedistributionRow};$s.Lines.Add($line)}
   $s.Confirmed=$true
  }else{$s | Add-Member -NotePropertyName ReferenceStatus -NotePropertyValue 'Октябрь рассчитан по текущему исходнику. Согласованные поправки прежнего файла не применялись.' -Force}
 }
 return $s
}
function Analyze-Files([string]$source,[string]$template,[string]$sheet,[datetime]$period,[bool]$useReference=$true){
 $source=[IO.Path]::GetFullPath($source);$template=[IO.Path]::GetFullPath($template)
 if($source -eq $template){throw 'Исходник и образец должны быть разными файлами.'}
 $excel=$null;$book=$null
 try{
  $excel=New-Excel;$book=$excel.Workbooks.Open($source,0,$true);$raw=Read-Matrix $book.Worksheets.Item($sheet);$book.Close($false);$book=$null
  $book=Open-Layout $excel $template $true;$templ=Read-Matrix $book.Worksheets.Item(1);$book.Close($false);$book=$null
 }finally{if($null -ne $book){$book.Close($false)};Close-Excel $excel}
 $s=Analyze-Matrix $raw $templ $source $template $sheet $period $useReference
 $s | Add-Member -NotePropertyName SourceFileHash -NotePropertyValue (File-Sha256 $source)
 $s | Add-Member -NotePropertyName TemplateFileHash -NotePropertyValue (File-Sha256 $template)
 return $s
}
function Validate-Session($s,[bool]$final){
 $errors=New-Object Collections.Generic.List[string]
 if($null -ne $s.Reference){
  if((Ledger-Hash $s) -ne $s.Reference.LedgerFingerprint){$errors.Add('Изменены исходные записи октябрьской сверки.')}
  if($s.Lines.Count -ne $s.Reference.Records.Count){$errors.Add('Изменён состав позиций октябрьской сверки.')}
  foreach($rec in $s.Reference.Records){
   $matched=@($s.Lines | Where-Object {$_.ReferenceRow -eq $rec.TargetRow})
   if($matched.Count -ne 1){$errors.Add("Отсутствует позиция образца $($rec.TargetRow).");continue}
   $l=$matched[0];for($i=0;$i -lt 7;$i++){if((Money $l.($script:Fields[$i])) -ne (Money $rec.Amounts[$i])){$errors.Add("Изменены суммы позиции образца $($rec.TargetRow).");break}}
   if($l.Name -cne $rec.Name -or $l.Contract -cne $rec.Contract -or $l.Supplier -cne $rec.Supplier -or $l.ContractDate -cne $rec.ContractDate -or $l.Number -cne $rec.Number){$errors.Add("Изменены реквизиты позиции образца $($rec.TargetRow).")}
  }
 }
 foreach($g in $s.Groups){
  $lines=@($s.Lines | Where-Object {$_.Group -eq $g.Id});$actual=[decimal]0;foreach($l in $lines){$actual+=Line-Total $l}
  $expected=[decimal]0
  if($null -ne $s.Reference){foreach($a in $g.Expected){$expected+=Money $a}}else{foreach($ev in $s.Events){if($ev.BlockId -eq $g.Id){$expected+=$ev.Amount}}}
  if($lines.Count -eq 0 -or $actual -ne $expected){$errors.Add("Договор $($g.Contract): сумма позиций $(Fixed $actual), ожидается $(Fixed $expected).")}
  if($null -eq $s.Reference){
   for($i=0;$i -lt 2;$i++){$sum=[decimal]0;foreach($l in $lines){$sum+=Money $l.($script:Fields[$i])};if($sum -ne (Money $g.Expected[$i])){$errors.Add("Договор $($g.Contract): изменена общая сумма отпуска / перераспределения.")}}
  }
 }
 if($s.BasisPath){foreach($l in $s.Lines){if([int]$l.BasisRow -lt 0){$errors.Add(('Договор '+$l.Contract+': выберите строку 1С или 0, чтобы не переносить сумму.'))}elseif([int]$l.BasisRow -gt 0 -and $s.BasisRows){$match=@($s.BasisRows | Where-Object {$_.Row -eq $l.BasisRow -and ($s.Automatic -or (Contract-Key $_.Contract) -eq (Contract-Key $l.Contract))});if($match.Count -ne 1){$errors.Add(('Договор '+$l.Contract+': в выбранной строке 1С нет этого договора.'))}}}}
 if($final -and -not $s.Automatic){foreach($l in $s.Lines){if(-not $l.Reviewed){$errors.Add("Не проверен договор $($l.Contract).")}elseif($null -eq $s.Reference -and -not (Text $l.Name)){$errors.Add('Есть позиция без наименования.')}};if($s.Warnings.Count -gt 0 -and -not $s.Confirmed){$errors.Add('Просмотрите записи с неопределённым периодом и поставьте флажок подтверждения.')}}
 return $errors.ToArray()
}
function Set-ExcelText($sheet,[int]$row,[int]$column,$value){$text=if($null -eq $value){''}else{[Convert]::ToString($value,$script:Invariant)};if($text -ne ''){$sheet.Cells.Item($row,$column).Formula="'$text"}}
function Set-ExcelAmount($sheet,[int]$row,[int]$column,$value){$amount=Money $value;if($amount -ne 0){$sheet.Cells.Item($row,$column).Formula=$amount.ToString('F2',$script:Invariant)}}
function Compact-Columns($sheet){
 $groups=@(@(2,2),@(3,6),@(7,10),@(11,13),@(14,15),@(16,17),@(18,20),@(21,22),@(23,24),@(25,27),@(28,29),@(30,32),@(33,34));$widths=@()
 foreach($g in $groups){$widths+=[double]$sheet.Range($sheet.Cells.Item(1,$g[0]),$sheet.Cells.Item(1,$g[1])).Width}
 $period=Text $sheet.Range('Q4').Value2;[void]$sheet.Range('Q4').MergeArea.UnMerge();[void]$sheet.Range('Q4').ClearContents();[void]$sheet.Range('P4:V4').Merge();Set-ExcelText $sheet 4 16 $period
 [void]$sheet.Columns.Item(35).Delete()
 for($i=$groups.Count-1;$i -ge 0;$i--){for($c=$groups[$i][1];$c -gt $groups[$i][0];$c--){[void]$sheet.Columns.Item($c).Delete()}}
 [void]$sheet.Columns.Item(1).Delete()
 for($i=0;$i -lt 13;$i++){
  $col=$sheet.Columns.Item($i+1);$target=$widths[$i];$lo=0.1;$hi=255.0;$best=0.0;$distance=1e9
  for($step=0;$step -lt 20;$step++){$candidate=($lo+$hi)/2;$col.ColumnWidth=$candidate;$actual=[double]$col.Width;$delta=[Math]::Abs($actual-$target);if($delta -lt $distance){$distance=$delta;$best=[double]$col.ColumnWidth};if($actual -lt $target){$lo=$candidate}else{$hi=$candidate}}
  $col.ColumnWidth=$best;if($distance -gt 0.76){throw "Не удалось сохранить ширину поля $($i+1)."}
 }
 return ,$widths
}
function Save-Report($s,[string]$output,[bool]$final=$false){
 if($s.BasisPath){return Save-MergedReport $s $output $final}
 $errors=@(Validate-Session $s $final);if($errors.Count){throw ($errors -join "`r`n")}
 $output=[IO.Path]::GetFullPath($output)
 if($output -eq [IO.Path]::GetFullPath($s.Source) -or $output -eq [IO.Path]::GetFullPath($s.Template)){throw 'Нельзя сохранять результат поверх исходника или образца.'}
 if([IO.Path]::GetExtension($output) -ne [IO.Path]::GetExtension($s.Template)){throw 'Сохраните результат в том же формате, что и образец.'}
 if($s.SourceFileHash -and (File-Sha256 $s.Source) -ne $s.SourceFileHash){throw 'Исходник изменился после чтения. Сохраните ручную работу отдельно и прочитайте исходник заново.'}
 if($s.TemplateFileHash -and (File-Sha256 $s.Template) -ne $s.TemplateFileHash){throw 'Образец изменился после чтения. Прочитайте файлы заново.'}
 $dir=[IO.Path]::GetDirectoryName($output);[void][IO.Directory]::CreateDirectory($dir)
 $temp=Join-Path $dir ('.report_'+[guid]::NewGuid().ToString('N')+[IO.Path]::GetExtension($output));Copy-Item -LiteralPath $s.Template -Destination $temp
 $e=$null;$book=$null
 try{
  $e=New-Excel;$book=Open-Layout $e $temp $false;$ws=$book.Worksheets.Item(1);$raw=Read-Matrix $ws
  if(-not (Text (Cell $raw 6 3)).Contains('лекарственного')){throw 'Изменена структура образца.'}
  $oldGrand=0;for($r=9;$r -le $raw.LastRow;$r++){if((Text (Cell $raw $r 14)).Replace(' ','') -eq 'ВСЕГО:'){$oldGrand=$r;break}}
  if($oldGrand -lt 10){throw 'В образце не найдена строка ВСЕГО.'}
  if($null -ne $s.Reference -and (Template-Hash $raw $s.Reference) -ne $s.TemplateFingerprint){throw 'Образец изменился после расчёта. Прочитайте файлы заново.'}
  $oldTotal=$oldGrand-1;$format=$ws.Cells.Item($oldTotal,16).NumberFormat
  $widths=Compact-Columns $ws
  [void]$ws.Range("A9:M$oldGrand").ClearContents()
  if($null -ne $s.Reference){$total=[int]$s.Reference.TotalRow;$grand=[int]$s.Reference.GrandRow}else{
   $count=$s.Lines.Count+@($s.Lines | Where-Object {$_.Redistribution -ne 0}).Count
   if($count -gt 65000){throw 'Слишком много строк для формата XLS.'}
   $donor=$book.Worksheets.Add();$donor.Name='__style';[void]$ws.Range('A10:M10').Copy($donor.Range('A1:M1'));[void]$donor.Range('A1:M1').ClearContents()
   [void]$ws.Rows.Item("9:$($oldTotal-1)").Delete();[void]$ws.Rows.Item("9:$($count+8)").Insert()
   $total=$count+9;$grand=$total+1
  }
  $row=9
  foreach($line in $s.Lines){
   if($null -ne $s.Reference){$row=[int]$line.ReferenceRow}else{[void]$donor.Range('A1:M1').Copy($ws.Range("A${row}:M${row}"));$ws.Rows.Item($row).RowHeight=20}
   Set-ExcelText $ws $row 1 $line.Number;Set-ExcelText $ws $row 2 $line.Name;Set-ExcelText $ws $row 3 $line.Supplier;Set-ExcelText $ws $row 4 $line.Contract;Set-ExcelText $ws $row 5 $line.ContractDate
   $cols=@(6,9,10,11,12,13);$fields=@('Transfer','Conversion','Fx','Commission','Discount','Other')
   for($i=0;$i -lt 6;$i++){Set-ExcelAmount $ws $row $cols[$i] $line.($fields[$i])}
   if($null -eq $s.Reference){$ws.Range("B${row}:E${row}").WrapText=$true;[void]$ws.Rows.Item($row).AutoFit()}
   $row++
   if($line.Redistribution -ne 0){
    if($null -ne $s.Reference){$rr=[int]$line.RedistributionRow;if($rr -le 0){throw 'Не задана строка перераспределения.'}}else{$rr=$row;[void]$donor.Range('A1:M1').Copy($ws.Range("A${rr}:M${rr}"));$ws.Rows.Item($rr).RowHeight=20;$row++}
    Set-ExcelText $ws $rr 5 'перераспр';Set-ExcelAmount $ws $rr 6 $line.Redistribution
   }
  }
  if($null -eq $s.Reference){$donor.Delete()}
  foreach($c in @('F','I','J','K','L','M')){$cell=$ws.Range("$c$total");$cell.Formula="=SUM(${c}9:${c}$($total-1))";$cell.MergeArea.NumberFormat=$format}
  $ws.Range("H$total").Formula="=IF(COUNT(H9:H$($total-1))=0,`"`",SUM(H9:H$($total-1)))";$ws.Range("H$total").NumberFormat=$format
  Set-ExcelText $ws $total 5 'ИТОГО:';Set-ExcelText $ws $grand 5 'ВСЕГО:'
  $ws.Range("F$grand").Formula="=SUM(F$total,I$total,J$total,K$total,L$total,M$total)";$ws.Range("F$grand").MergeArea.NumberFormat=$format
  $period=[datetime]::ParseExact($s.PeriodKey,'yyyy-MM',$script:Invariant);Set-ExcelText $ws 4 6 ('за период с '+$period.ToString('dd.MM.yyyy')+' по '+$period.AddMonths(1).AddDays(-1).ToString('dd.MM.yyyy'))
  $e.Calculation=-4105;$e.CalculateFull();$actual=Money $ws.Range("F$grand").Value2;$expected=Session-Total $s
  if($actual -ne $expected){throw "Не сошёлся итог Excel: $actual вместо $expected."}
  foreach($c in @('F','I','J','K','L','M')){if((Text $ws.Range("$c$total").Text).Contains('#')){throw "Не помещается итог в поле $c."}}
  for($i=0;$i -lt 13;$i++){if([Math]::Abs([double]$ws.Columns.Item($i+1).Width-$widths[$i]) -gt 0.76){throw 'Изменилась ширина поля.'}}
  $ws.PageSetup.PrintArea='$A$3:$M$'+$grand;[void]$ws.Activate();[void]$ws.Range('A1').Select();$book.Save();$book.Close($false);$book=$null
  Copy-Item -LiteralPath $temp -Destination $output -Force
 }finally{if($null -ne $book){$book.Close($false)};Close-Excel $e;if(Test-Path -LiteralPath $temp){Remove-Item -LiteralPath $temp}}
 $auditFolder=Join-Path $dir 'Проверки';[void][IO.Directory]::CreateDirectory($auditFolder)
 $audit=Join-Path $auditFolder ([IO.Path]::GetFileNameWithoutExtension($output)+'.html');Write-Audit $s $audit
 return $output
}
function Html($value){return [Net.WebUtility]::HtmlEncode((Text $value))}
function Write-Audit($s,[string]$path){
 $b=New-Object Text.StringBuilder
 [void]$b.Append("<!doctype html><html lang='ru'><meta charset='utf-8'><title>Проверка отчёта</title><style>body{font:15px Segoe UI,Arial;margin:30px}table{border-collapse:collapse;margin:18px 0}td,th{border:1px solid #ccc;padding:6px;text-align:left}th{background:#eee}h1{font-size:24px}</style><h1>Проверка отчёта $(Html $s.PeriodKey)</h1><p>Исходник: $(Html $s.Source)<br>Вкладка: $(Html $s.Sheet)<br>Образец: $(Html $s.Template)</p><p>Месяц определяется только по колонке 22. Суммы записей — из колонки 23. Даты других колонок на отбор не влияют. Колонки 17–21 используются для расшифровки расходов. В результате 13 рабочих колонок; ширина полей сохранена. Все нижние итоги — формулы Excel; учреждения не прибавляются к ВСЕГО повторно.</p>")
 $source=Source-Total $s;$total=Session-Total $s
 [void]$b.Append("<p>Сумма выбранных записей: $(Fixed $source). Сумма результата: $(Fixed $total). Разница: $(Fixed ($total-$source)).</p>")
 if($null -ne $s.Reference){[void]$b.Append('<p>Применена сверка с октябрьским образцом при совпадении контрольных отпечатков. Разбивки нескольких препаратов и пять согласованных отличий взяты из образца, а не вычислены независимо из исходника. Ручное ВСЕГО образца не используется.</p><table><tr><th>Договор</th><th>Основание</th></tr>');foreach($g in $s.Reference.Groups){if($g.Reason -or $g.SplitFromReference){[void]$b.Append("<tr><td>$(Html $g.Contract)</td><td>$(Html $g.Reason) $(if($g.SplitFromReference){'Разбивка по препаратам из образца.'})</td></tr>")}};[void]$b.Append('</table>')}
 [void]$b.Append('<h2>Позиции</h2><table><tr><th>Препарат</th><th>Договор</th><th>Отпуск</th><th>Перераспр.</th><th>Валюта</th><th>Курс. разн.</th><th>Комиссия</th><th>Скидки</th><th>Прочие</th><th>Проверено</th></tr>')
 foreach($l in $s.Lines){[void]$b.Append("<tr><td>$(Html $l.Name)</td><td>$(Html $l.Contract)</td>");foreach($f in $script:Fields){[void]$b.Append("<td>$(Fixed $l.$f)</td>")};[void]$b.Append("<td>$($l.Reviewed)</td></tr>")};[void]$b.Append('</table><h2>Месячные записи исходника</h2><table><tr><th>Строка</th><th>Договор</th><th>Колонка 22</th><th>Колонка 23</th></tr>')
 foreach($ev in $s.Events){[void]$b.Append("<tr><td>$($ev.Row)</td><td>$(Html $ev.Contract)</td><td>$(Html $ev.Label)</td><td>$(Fixed $ev.Amount)</td></tr>")};[void]$b.Append('</table><h2>Замечания</h2><ul>');foreach($g in $s.Groups){if($g.Note){[void]$b.Append("<li>$(Html $g.Contract): $(Html $g.Note)</li>")}};foreach($w in $s.Warnings){[void]$b.Append("<li>$(Html $w)</li>")};[void]$b.Append('</ul></html>')
 [IO.File]::WriteAllText($path,$b.ToString(),(New-Object Text.UTF8Encoding($true)))
}

# Two-input reports: the 1C book owns all rows and metadata; the ledger supplies explicit replacements.
function Contract-Key($value){$key=[regex]::Replace((Text $value).ToUpperInvariant(),'\s','').TrimStart([char]'№');if((Normal $key) -match '^дистриб.*контр') {return '@DISTRIBUTOR'};return $key}
function Read-Basis($matrix,[datetime]$period){
 if(-not (Text (Cell $matrix 6 3)).Contains('лекарственного') -or (Text (Cell $matrix 7 11)) -ne 'номер'){throw 'В файле 1С не найдена структура отчёта МЗ РБ (заголовки строк 6–8).'}
 $date=Parse-Date (Cell $matrix 4 17)
 if($null -eq $date -or $date.Year -ne $period.Year -or $date.Month -ne $period.Month){throw 'Период в заголовке отчёта 1С не совпадает с выбранным месяцем. Выберите подходящий файл или вкладку.'}
 $grand=0
 for($r=9;$r -le $matrix.LastRow;$r++){if((Text (Cell $matrix $r 14)).Replace(' ','') -eq 'ВСЕГО:'){if($grand){throw 'В отчёте 1С несколько строк ВСЕГО.'};$grand=$r}}
 if($grand -lt 11){throw 'В файле 1С не найдена строка ВСЕГО после данных.'}
 $columns=@(2,3,7,11,14,16,18,21,23,25,28,30,33)
 $rows=New-Object Collections.Generic.List[object];$main=New-Object Collections.Generic.List[object];$last=$null
 for($r=9;$r -lt ($grand-1);$r++){
  for($c=$matrix.FirstColumn;$c -le $matrix.LastColumn;$c++){if($c -notin $columns -and (Text (Cell $matrix $r $c)) -ne ''){throw "Файл 1С: в строке $r, колонке $c есть данные вне 13 полей. Удаление этой колонки привело бы к потере данных."}}
  $values=New-Object 'object[]' 13
  for($i=0;$i -lt 13;$i++){$values[$i]=Cell $matrix $r $columns[$i]}
  foreach($i in @(5,7,8,9,10,11,12)){if((Text $values[$i]) -ne ''){$values[$i]=Money $values[$i]}}
  $d=Parse-Date $values[4];if($null -ne $d){$values[4]=$d.ToString('dd.MM.yyyy')}
  $item=[pscustomobject]@{Row=$r;Values=$values;ParentRow=0;IsRedistribution=$false}
  if((Text $values[0]) -ne '' -and (Text $values[3]) -ne ''){$last=$item;$main.Add($item)}
  elseif($null -ne $last){$item.ParentRow=$last.Row;$item.IsRedistribution=(Normal $values[4]).StartsWith('перераспр')}
  $rows.Add($item)
 }
 foreach($item in $rows){if((Normal $item.Values[0]) -eq 'итого' -or (Normal $item.Values[4]) -eq 'итого') {throw 'Внутри данных 1С найдены промежуточные итоги. Нужна отдельная проверка структуры, чтобы не сложить суммы дважды.'}}
 if($main.Count -eq 0){throw 'В отчёте 1С не найдены строки договоров.'}
 return [pscustomobject]@{Rows=$rows;Main=$main;TotalRow=$grand-1;GrandRow=$grand}
}
function Set-BasisMatches($s,$basis){
 foreach($line in $s.Lines){
  $contract=@($basis.Main | Where-Object {(Contract-Key $_.Values[3]) -eq (Contract-Key $line.Contract)})
  $exact=@($contract | Where-Object {
   (Normal $_.Values[1]) -eq (Normal $line.Name) -and
   (-not (Text $_.Values[2]) -or -not $line.Supplier -or (Normal $_.Values[2]) -eq (Normal $line.Supplier)) -and
   (-not (Text $_.Values[4]) -or -not $line.ContractDate -or (Text $_.Values[4]) -eq $line.ContractDate)
  })
  $row=if($exact.Count -eq 1){$exact[0].Row}else{-1}
  $line | Add-Member NoteProperty BasisRow ([int]$row) -Force
  if($row -lt 0){
   foreach($other in $s.Lines){if($other.Group -eq $line.Group){$other.Reviewed=$false}}
   $g=@($s.Groups | Where-Object {$_.Id -eq $line.Group})[0];$g.NeedsReview=$true
   $candidates=($contract | ForEach-Object {'строка '+$_.Row+': '+(Text $_.Values[1])+' / '+(Text $_.Values[2])}) -join '; '
   $g.Note+=' Выберите строку в 1С для «'+$line.Name+'» (0 — не переносить эту сумму). Кандидаты: '+$candidates+'.'
  }
 }
}
function Overlay-Fields($s,$group){
 if($group.PSObject.Properties['AutoFields']){return @($group.AutoFields)}
 $categories=@($s.Events | Where-Object {$_.BlockId -eq $group.Id} | ForEach-Object {$_.Category})
 $fields=New-Object Collections.Generic.List[string]
 $map=@{Transfer=@('Transfer');Redistribution=@('Redistribution');Commission=@('Commission');Fx=@('Fx');Discount=@('Discount');Expenses=@('Conversion','Other');Other=@('Other');Vat=@('Other');Unknown=@('Other')}
 foreach($category in $categories){foreach($f in $map[$category]){if(-not $fields.Contains($f)){$fields.Add($f)}}}
 # Reconciled or manually allocated nonzero components are also intentional replacements.
 foreach($f in $script:Fields){if(@($s.Lines | Where-Object {$_.Group -eq $group.Id -and (Money $_.$f) -ne 0}).Count -gt 0 -and -not $fields.Contains($f)){$fields.Add($f)}}
 return $fields.ToArray()
}
function Merge-Basis($s,$basis){
 if($s.Automatic){return Merge-AutomaticContracts $s $basis}
 $rows=New-Object Collections.Generic.List[object];$byRow=@{}
 foreach($r in $basis.Rows){$copy=[pscustomobject]@{Row=$r.Row;Values=[object[]]$r.Values.Clone();ParentRow=$r.ParentRow;IsRedistribution=$r.IsRedistribution};$rows.Add($copy);$byRow[[int]$r.Row]=$copy}
 $targets=@{};$changes=New-Object Collections.Generic.List[object];$skipped=New-Object Collections.Generic.List[object]
 foreach($g in $s.Groups){
  $fields=@(Overlay-Fields $s $g)
  foreach($line in @($s.Lines | Where-Object {$_.Group -eq $g.Id})){
   $row=[int]$line.BasisRow
   if($row -lt 0){throw ('Договор '+$line.Contract+': выберите строку 1С для «'+$line.Name+'» или 0, чтобы не переносить сумму.')}
   if($row -gt 0){
    $main=@($basis.Main | Where-Object {$_.Row -eq $row})
    $contractOK=$main.Count -eq 1 -and (Contract-Key $main[0].Values[3]) -eq (Contract-Key $line.Contract)

    if(-not $contractOK){throw ('Договор '+$line.Contract+': выбранная строка '+$row+' в 1С относится к другому договору или не является позицией.')}
    $key='row:'+$row;$target=$byRow[$row]
   }else{$skipped.Add($line);continue}
   if(-not $targets.ContainsKey($key)){$targets[$key]=[pscustomobject]@{Target=$target;Amounts=@{}}}
   foreach($f in $fields){if(-not $targets[$key].Amounts.ContainsKey($f)){$targets[$key].Amounts[$f]=[decimal]0};$targets[$key].Amounts[$f]+=Money $line.$f}
  }
 }
 $indices=@{Transfer=5;Conversion=8;Fx=9;Commission=10;Discount=11;Other=12}
 foreach($t in $targets.Values){
  foreach($f in @('Transfer','Conversion','Fx','Commission','Discount','Other','Redistribution')){
   if(-not $t.Amounts.ContainsKey($f)){continue}
   $target=$t.Target
   if($f -eq 'Redistribution'){
    $existing=@($rows | Where-Object {$target.Row -gt 0 -and $_.ParentRow -eq $target.Row -and $_.IsRedistribution})
    if($existing.Count -gt 1){throw ('Несколько строк перераспределения в 1С для строки '+$target.Row+'. Требуется уточнить распределение.')}
    if($existing.Count){$target=$existing[0]}else{$t.Amounts[$f]=(Money $target.Values[5])+$t.Amounts[$f]}
    $i=5
   }else{$i=$indices[$f]}
   $changes.Add([pscustomobject]@{BasisRow=$t.Target.Row;Contract=$t.Target.Values[3];Name=$t.Target.Values[1];Field=$f;Before=$target.Values[$i];After=$t.Amounts[$f]})
   $target.Values[$i]=$t.Amounts[$f]
  }
 }
 $total=[decimal]0;foreach($r in $rows){foreach($i in @(5,8,9,10,11,12)){$total+=Money $r.Values[$i]}}
 return [pscustomobject]@{Rows=$rows;Total=$total;Changes=$changes;Skipped=$skipped;PositionCount=@($rows | Where-Object {(Text $_.Values[0]) -ne ''}).Count}
}
function Analyze-Inputs([string]$source,[string]$sheet,[datetime]$period,[string]$basisPath,[string]$basisSheet,[bool]$fresh=$false){
 $source=[IO.Path]::GetFullPath($source);$basisPath=[IO.Path]::GetFullPath($basisPath)
 if($source -eq $basisPath){throw 'Выберите два разных файла: расчёты и отчёт из 1С.'}
 $cache=Cache-Path $source $sheet $period.ToString('yyyy-MM') $basisPath $basisSheet
 if(-not $fresh -and (Test-Path -LiteralPath $cache)){try{$saved=Restore-Session $cache;if(Session-Current $saved $source $sheet $period.ToString('yyyy-MM') $basisPath $basisPath $basisSheet){return $saved}}catch{}}
 $sourceHash=File-Sha256 $source;$basisHash=File-Sha256 $basisPath;$excel=$null;$book=$null
 try{
  $excel=New-Excel;$book=$excel.Workbooks.Open($source,0,$true);$raw=Read-Matrix $book.Worksheets.Item($sheet);$book.Close($false);$book=$null
  $book=$excel.Workbooks.Open($basisPath,0,$true);$matrix=Read-Matrix $book.Worksheets.Item($basisSheet);$basis=Read-Basis $matrix $period;$book.Close($false);$book=$null
 }finally{if($null -ne $book){$book.Close($false)};Close-Excel $excel}
 if($sourceHash -ne (File-Sha256 $source) -or $basisHash -ne (File-Sha256 $basisPath)){throw 'Один из файлов изменился во время чтения. Повторите формирование.'}
 $s=Analyze-Matrix $raw $matrix $source $basisPath $sheet $period $false $true $true
 foreach($p in @{BasisPath=$basisPath;BasisSheet=$basisSheet;SourceFileHash=$sourceHash;TemplateFileHash=$basisHash;BasisFileHash=$basisHash;EngineHash=(File-Sha256 (Join-Path $PSScriptRoot 'Core.ps1'))}.GetEnumerator()){$s | Add-Member NoteProperty $p.Key $p.Value -Force}
 $s | Add-Member NoteProperty BasisRows @($basis.Main | ForEach-Object {[pscustomobject]@{Row=$_.Row;Contract=$_.Values[3]}}) -Force
 Set-AutomaticBasisMatches $s $basis;Save-Cache $s;return $s
}
function Save-MergedReport($s,[string]$output,[bool]$final){
 $errors=@(Validate-Session $s $final);if($errors.Count){throw ($errors -join "`r`n")}
 $output=[IO.Path]::GetFullPath($output)
 if($output -eq [IO.Path]::GetFullPath($s.Source) -or $output -eq [IO.Path]::GetFullPath($s.BasisPath)){throw 'Нельзя сохранять результат поверх входных файлов.'}
 if([IO.Path]::GetExtension($output) -ne '.xls'){throw 'Сохраните результат в формате XLS.'}
 if((File-Sha256 $s.Source) -ne $s.SourceFileHash -or (File-Sha256 $s.BasisPath) -ne $s.BasisFileHash){throw 'Один из входных файлов изменился после чтения. Прочитайте оба файла заново.'}
 $dir=[IO.Path]::GetDirectoryName($output);[void][IO.Directory]::CreateDirectory($dir)
 $temp=Join-Path $dir ('.report_'+[guid]::NewGuid().ToString('N')+[IO.Path]::GetExtension($s.BasisPath));Copy-Item -LiteralPath $s.BasisPath -Destination $temp
 $excel=$null;$book=$null
 if((File-Sha256 $temp) -ne $s.BasisFileHash){Remove-Item -LiteralPath $temp;throw 'Файл 1С изменился во время копирования. Повторите формирование.'}
 try{
  $excel=New-Excel;$book=$excel.Workbooks.Open($temp,0,$false);$ws=$book.Worksheets.Item($s.BasisSheet)
  $basis=Read-Basis (Read-Matrix $ws) ([datetime]::ParseExact($s.PeriodKey,'yyyy-MM',$script:Invariant));$plan=Merge-Basis $s $basis
  if($plan.Rows.Count -gt 65000){throw 'Слишком много строк для формата XLS.'}
  $widths=Compact-Columns $ws;$format=$ws.Cells.Item($basis.TotalRow,6).NumberFormat
  $donor=$book.Worksheets.Add();[void]$ws.Range("A9:M$($basis.TotalRow-1)").Copy($donor.Range('A1'));[void]$donor.UsedRange.ClearContents()
  [void]$ws.Rows.Item("9:$($basis.TotalRow-1)").Delete();[void]$ws.Rows.Item("9:$($plan.Rows.Count+8)").Insert()
  $row=9
  foreach($item in $plan.Rows){
   $styleRow=if($item.Row -gt 0){$item.Row-8}else{1}
   [void]$donor.Range("A${styleRow}:M${styleRow}").Copy($ws.Range("A${row}:M${row}"));$ws.Rows.Item($row).RowHeight=$donor.Rows.Item($styleRow).RowHeight
   for($i=0;$i -lt 13;$i++){
    $v=$item.Values[$i];if((Text $v) -eq ''){continue}
    if($i -in @(5,7,8,9,10,11,12)){$ws.Cells.Item($row,$i+1).Formula=(Money $v).ToString('F2',$script:Invariant)}else{Set-ExcelText $ws $row ($i+1) $v}
   }
   $row++
  }
  $donor.Delete();$total=$row;$grand=$row+1
  [void]$ws.Range("A${total}:M${grand}").ClearContents()
  foreach($c in @('F','H','I','J','K','L','M')){$ws.Range("$c$total").Formula="=SUM(${c}9:${c}$($total-1))";$ws.Range("$c$total").NumberFormat=$format}
  Set-ExcelText $ws $total 5 'ИТОГО:';Set-ExcelText $ws $grand 5 'ВСЕГО:'
  $ws.Range("F$grand").Formula="=SUM(F$total,I$total,J$total,K$total,L$total,M$total)";$ws.Range("F$grand").MergeArea.NumberFormat=$format
  $excel.Calculation=-4105;$excel.CalculateFull()
  if((Money $ws.Range("F$grand").Value2) -ne $plan.Total){throw 'Итог Excel не совпал с суммой объединённых данных.'}
  for($i=0;$i -lt 13;$i++){if([Math]::Abs([double]$ws.Columns.Item($i+1).Width-$widths[$i]) -gt 0.76){throw 'Изменилась ширина поля.'}}
  $ws.PageSetup.PrintArea='$A$3:$M$'+$grand
  # Only the selected report sheet is exported, in a real XLS container even for XLSX input.
  [void]$ws.Copy();$result=$excel.ActiveWorkbook
  try{$result.CheckCompatibility=$false;$result.SaveAs($output,56)}finally{$result.Close($false)}
  if($s.Automatic){$s | Add-Member NoteProperty AutomationNotes @($plan.Notes) -Force}
  $s | Add-Member NoteProperty ResultTotal $plan.Total -Force;$s | Add-Member NoteProperty ResultPositions $plan.PositionCount -Force
  $s | Add-Member NoteProperty ResultChanges $plan.Changes.Count -Force
  $s | Add-Member NoteProperty ResultNotes (@($s.AutomationNotes | Where-Object {$_.Relevant}).Count+@($s.Warnings).Count) -Force
  $auditDir=Join-Path $dir 'Проверки';[void][IO.Directory]::CreateDirectory($auditDir)
  $html=New-Object Text.StringBuilder
  [void]$html.Append("<!doctype html><meta charset='utf-8'><title>Проверка объединения</title><h1>Объединение отчёта $(Html $s.PeriodKey)</h1><p>Основа 1С: $(Html $s.BasisPath), вкладка $(Html $s.BasisSheet)<br>Расчёты: $(Html $s.Source), вкладка $(Html $s.Sheet)</p><p>Явный ноль заменяет значение. Отсутствующая категория сохраняет значение 1С. Учреждения не прибавляются повторно к ВСЕГО.</p><table border='1'><tr><th>Строка 1С</th><th>Договор</th><th>Препарат</th><th>Поле</th><th>Было</th><th>Стало</th></tr>")
  $labels=@{Transfer='Отпуск';Redistribution='Перераспределение в отпуске';Conversion='Покупка валюты';Fx='Курсовые разницы';Commission='Комиссия';Discount='Скидки';Other='Прочие расходы';Name='Наименование'}
  foreach($c in $plan.Changes){[void]$html.Append("<tr><td>$($c.BasisRow)</td><td>$(Html $c.Contract)</td><td>$(Html $c.Name)</td><td>$(Html $labels[$c.Field])</td><td>$(Html $c.Before)</td><td>$(Html $(if($c.Field -eq 'Name'){$c.After}else{Fixed $c.After}))</td></tr>")};[void]$html.Append('</table><h2>Не перенесены: значения 1С сохранены</h2>');foreach($line in $plan.Skipped){[void]$html.Append("<p>$(Html $line.Contract) — $(Html $line.Name): $(Html (($script:Fields | ForEach-Object {$labels[$_]+': '+(Fixed $line.$_)}) -join '; '))</p>")};[void]$html.Append("<p>ВСЕГО: $(Html (Fixed $plan.Total))</p>")
  [void]$html.Append('<h2>Замечания автоматического сопоставления</h2>');foreach($note in $s.AutomationNotes){[void]$html.Append("<p>Строка расчётов $($note.SourceRow), $(Html $note.Contract), $(Html $note.Name): $(Html $note.Reason)</p>")};foreach($warning in $s.Warnings){[void]$html.Append("<p>$(Html $warning)</p>")}
  [IO.File]::WriteAllText((Join-Path $auditDir ([IO.Path]::GetFileNameWithoutExtension($output)+'.html')),$html.ToString(),[Text.Encoding]::UTF8)
 }finally{if($null -ne $book){$book.Close($false)};Close-Excel $excel;if(Test-Path -LiteralPath $temp){Remove-Item -LiteralPath $temp}}
 return $output
}

function Source-DateKey($value){$d=Parse-Date $value;if($null -eq $d){return ''};return $d.ToString('yyyy-MM-dd')}
function Basis-ContractSets($basis){
 $sets=[ordered]@{}
 foreach($row in $basis.Main){
  $contract=Contract-Key $row.Values[3]
  $key=$contract+'|'+(Source-DateKey $row.Values[4])+'|'+(Normal $row.Values[2])
  if(-not $sets.Contains($key)){$sets[$key]=[pscustomobject]@{Key=$key;Contract=$contract;Date=(Source-DateKey $row.Values[4]);Supplier=(Normal $row.Values[2]);Rows=(New-Object Collections.Generic.List[object])}}
  $sets[$key].Rows.Add($row)
 }
 return @($sets.Values)
}
function Match-ContractSet($sets,$s,$group){
 $candidates=@($sets | Where-Object {$_.Contract -eq (Contract-Key $group.Contract)})
 if($candidates.Count -le 1){return $candidates}
 $line=@($s.Lines | Where-Object {$_.Group -eq $group.Id})[0]
 if($null -eq $line){$line=$group}
 $date=Source-DateKey $line.ContractDate
 if($date){$dated=@($candidates | Where-Object {$_.Date -eq $date});if($dated.Count){$candidates=$dated}}
 if($candidates.Count -gt 1){
  $supplier=Normal $line.Supplier
  if($supplier){$named=@($candidates | Where-Object {$_.Supplier -and ($supplier.Contains($_.Supplier) -or $_.Supplier.Contains($supplier))});if($named.Count){$candidates=$named}}
 }
 return $candidates
}
function Set-AutomaticBasisMatches($s,$basis){
 # Mapping and field allocation happen together against the current basis at export.
 foreach($line in $s.Lines){$line | Add-Member NoteProperty BasisRow 0 -Force}
 $s | Add-Member NoteProperty Automatic $true -Force
 $s | Add-Member NoteProperty AutomationNotes @() -Force
 $s | Add-Member NoteProperty ZeroMissingFields $true -Force
}
function Set-ContractValue($changes,$row,[int]$index,[string]$field,$value,$sourceRows){
 if($field -eq 'Name'){$same=(Text $row.Values[$index]) -ceq (Text $value)}else{$same=(Money $row.Values[$index]) -eq (Money $value)}
 if($same){return}
 $changes.Add([pscustomobject]@{BasisRow=$row.Row;Contract=$row.Values[3];Name=$row.Values[1];Field=$field;Before=$row.Values[$index];After=$value;SourceRows=($sourceRows -join ', ')})
 $row.Values[$index]=$value
}
function Merge-AutomaticContracts($s,$basis){
 $rows=New-Object Collections.Generic.List[object];$byRow=@{}
 foreach($r in $basis.Rows){$copy=[pscustomobject]@{Row=$r.Row;Values=[object[]]$r.Values.Clone();ParentRow=$r.ParentRow;IsRedistribution=$r.IsRedistribution};$rows.Add($copy);$byRow[[int]$r.Row]=$copy}
 $sets=@(Basis-ContractSets $basis);$incoming=@{};$changes=New-Object Collections.Generic.List[object];$notes=New-Object Collections.Generic.List[object];$skipped=New-Object Collections.Generic.List[object]
 $groups=New-Object Collections.Generic.List[object];foreach($g in $s.Groups){$groups.Add($g)}
 if($s.ZeroMissingFields){
  $activeIds=@($s.Groups | ForEach-Object {$_.Id})
  foreach($contract in $s.AllContracts){if($contract.Id -notin $activeIds){$groups.Add([pscustomobject]@{Id=$contract.Id;Row=$contract.Row;Contract=$contract.Contract;ContractDate=$contract.ContractDate;Supplier=$contract.Supplier;Expected=@(0,0,0,0,0,0,0)})}}
 }
 foreach($g in $groups){
  $match=@(Match-ContractSet $sets $s $g)
  if($match.Count -ne 1){
   $reason=if($match.Count -eq 0){'В 1С нет этого договора; новые позиции не добавлялись.'}else{'Номер договора повторяется у разных поставщиков/дат; значения 1С сохранены.'}
   $notes.Add([pscustomobject]@{SourceRow=$g.Row;Contract=$g.Contract;Name='';Reason=$reason;Relevant=($match.Count -gt 1)})
   foreach($line in @($s.Lines | Where-Object {$_.Group -eq $g.Id})){$skipped.Add($line)}
   continue
  }
  $key=$match[0].Key
  if(-not $incoming.ContainsKey($key)){$incoming[$key]=[pscustomobject]@{Amounts=@{};Groups=(New-Object Collections.Generic.List[object]);Fields=(New-Object Collections.Generic.List[string])}}
  $incoming[$key].Groups.Add($g)
  $fields=if($s.ZeroMissingFields){$script:Fields}else{@(Overlay-Fields $s $g)}
  foreach($field in $fields){if(-not $incoming[$key].Fields.Contains($field)){$incoming[$key].Fields.Add($field)}}
  for($i=0;$i -lt $script:Fields.Count;$i++){$f=$script:Fields[$i];if(-not $incoming[$key].Amounts.ContainsKey($f)){$incoming[$key].Amounts[$f]=[decimal]0};$incoming[$key].Amounts[$f]+=Money $g.Expected[$i]}
 }
 $indices=@{Transfer=5;Conversion=8;Fx=9;Commission=10;Discount=11;Other=12};$matched=0
 # Output order and membership are driven exclusively by the 1C rows.
 foreach($set in $sets){
  if(-not $incoming.ContainsKey($set.Key)){continue};$matched++
  $source=$incoming[$set.Key];$sources=@($source.Groups | ForEach-Object {$_.Row})
  $members=@($set.Rows | ForEach-Object {$byRow[[int]$_.Row]})
  $named=@($members | Where-Object {(Text $_.Values[1]) -ne ''})
  $active=if($named.Count){$named}else{$members}
  $primary=$active[0]
  # A blank name may be completed from one unambiguous source name, without adding a row.
  if($members.Count -eq 1 -and -not (Text $primary.Values[1])){
   $ids=@($source.Groups | ForEach-Object {$_.Id});$names=@($s.Lines | Where-Object {$_.Group -in $ids -and $_.Name -ne 'Укажите наименование'} | ForEach-Object {$_.Name} | Select-Object -Unique)
   if($names.Count -eq 1){Set-ContractValue $changes $primary 1 'Name' $names[0] $sources}
  }
  foreach($field in @('Transfer','Conversion','Fx','Commission','Discount','Other')){
   if(-not $source.Fields.Contains($field)){continue}
   $index=$indices[$field];$wanted=Money $source.Amounts[$field];$current=[decimal]0
   foreach($member in $members){$current+=Money $member.Values[$index]}
   if($current -eq $wanted -and $wanted -ne 0){continue} # Preserve a correct existing split across drugs.
   if($active.Count -gt 1 -and $field -eq 'Transfer' -and $wanted -ne 0){
    $notes.Add([pscustomobject]@{SourceRow=($sources -join ', ');Contract=$primary.Values[3];Name='';Reason='Общий отпуск договора не совпал с разбивкой 1С по препаратам. Разбивка сохранена; произвольное распределение не выполнялось.';Relevant=$true});continue
   }
   # Contract-level costs are recorded once in the first populated 1C position.
   # Do not repeat the full contract amount for each drug.
   foreach($member in $members){$value=if($member.Row -eq $primary.Row){$wanted}else{[decimal]0};Set-ContractValue $changes $member $index $field $value $sources}
  }
  if($source.Fields.Contains('Redistribution')){
   $parentRows=@($members | ForEach-Object {$_.Row});$redis=@($rows | Where-Object {$_.IsRedistribution -and $_.ParentRow -in $parentRows})
   $wanted=Money $source.Amounts['Redistribution']
   if($redis.Count){$current=[decimal]0;foreach($r in $redis){$current+=Money $r.Values[5]};if($current -ne $wanted){for($i=0;$i -lt $redis.Count;$i++){Set-ContractValue $changes $redis[$i] 5 'Redistribution' $(if($i -eq 0){$wanted}else{[decimal]0}) $sources}}}
   elseif($wanted -ne 0){Set-ContractValue $changes $primary 5 'Redistribution' ((Money $primary.Values[5])+$wanted) $sources}
  }
 }
 $total=[decimal]0;foreach($r in $rows){foreach($i in @(5,8,9,10,11,12)){$total+=Money $r.Values[$i]}}
 return [pscustomobject]@{Rows=$rows;Changes=$changes;Skipped=$skipped;Notes=$notes;MatchedContracts=$matched;Total=$total;PositionCount=$basis.Main.Count}
}
