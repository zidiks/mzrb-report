$ErrorActionPreference='Stop'
$script:Invariant=[Globalization.CultureInfo]::InvariantCulture
$script:Fields=@('Transfer','Redistribution','Conversion','Fx','Commission','Discount','Other')
$script:Months=@('январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь')
function Text($v) { if($null -eq $v){return ''};return [Convert]::ToString($v,$script:Invariant).Trim() }
function Money($v) {
 if($null -eq $v -or (Text $v) -eq ''){return [decimal]0}
 $str=(Text $v).Replace([string][char]160,'').Replace(' ','').Replace(',','.')
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
function Analyze-Matrix($raw,$template,[string]$source,[string]$templatePath,[string]$sheet,[datetime]$period,[bool]$useReference=$true){
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
 foreach($block in $blocks){
  $events=New-Object Collections.Generic.List[object]
  foreach($r in $block.Rows){
   $label=Text (Cell $raw $r 22);$amount=Money (Cell $raw $r 23)
   if($label -eq ''){if($amount -ne 0){$s.Warnings.Add("Строка $r, $($block.Contract): сумма $amount без периода; не включена.")};continue}
   $type=Label-Type (Cell $raw $r 22)
   if($type.Month -eq 0){$s.Warnings.Add("Строка $r, $($block.Contract): месяц метки «$label» не определён, сумма $amount не включена.");continue}
   if($type.Month -ne $period.Month -or ($type.Year -ne 0 -and $type.Year -ne $period.Year)){continue}
   $event=[pscustomobject]@{BlockId=$block.Id;Row=$r;Contract=$block.Contract;Label=$label;Amount=$amount;Category=$type.Category};$events.Add($event);$s.Events.Add($event)
  }
  if($events.Count -eq 0){continue}
  $supplier=New-Object Collections.Generic.List[string];$names=New-Object Collections.Generic.List[string];$costs=New-Object Collections.Generic.List[object]
  foreach($r in $block.Rows){
   $sup=Text (Cell $raw $r 8);if($sup){$supplier.Add($sup)}
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
  $known=@();$key=Normal $block.Contract;if($refs.ContainsKey($key)){$known=@($refs[$key].ToArray())}
  if($names.Count -gt 1){$joined=($names -join ' ').Replace('- ','-');$match=@($known | Where-Object {(Normal $_) -eq (Normal $joined)});$continuation=$true;for($i=1;$i -lt $names.Count;$i++){if($names[$i].Length -gt 2 -and -not [char]::IsLower($names[$i][0])){$continuation=$false}}
   if($match.Count -gt 0 -or $continuation -or @($names | Select-Object -SkipLast 1 | Where-Object {$_.EndsWith('-')}).Count -gt 0){$names.Clear();$names.Add($(if($match.Count){$match[0]}else{$joined}))}
  }
  if($names.Count -eq 1 -and $known.Count -gt 1){$names.Clear();foreach($n in $known){$names.Add($n)};$review=$true}
  if($names.Count -eq 0){$names.Add('Укажите наименование');$review=$true}
  if($names.Count -gt 1){$review=$true;$notes.Add('Несколько препаратов: общая сумма временно стоит в первой строке. Распределите её, не увеличивая сумму договора.')}
  if($null -eq (Parse-Date $block.ContractDate)){$review=$true;$notes.Add('Проверьте дату договора в строке под его номером.')}
  $vat=Sum-Events $events 'Vat';if($vat -ne 0){$review=$true;$notes.Add('НДС включён в прочие расходы; требуется проверка.')}
  if(@($events | Where-Object {$_.Category -eq 'Unknown'}).Count){$review=$true;$notes.Add('Есть нераспознанная месячная метка: распределите сумму вручную.')}
  $expected=@((Sum-Events $events 'Transfer'),(Sum-Events $events 'Redistribution'),$conversion,(Sum-Events $events 'Fx'),$commission,(Sum-Events $events 'Discount'),((Sum-Events $events 'Commission')+(Sum-Events $events 'Expenses')+$vat+(Sum-Events $events 'Other')-$conversion-$commission))
  $g=[pscustomobject]@{Id=$block.Id;Row=$block.Row;Contract=$block.Contract;Expected=$expected;NeedsReview=$review;Note=($notes -join ' ')};$s.Groups.Add($g)
  for($i=0;$i -lt $names.Count;$i++){$line=New-Line $block $names[$i];$line.Note=$g.Note;$line.Reviewed=-not $review;if($i -eq 0){for($j=0;$j -lt 7;$j++){$line.($script:Fields[$j])=[decimal]$expected[$j]}};$s.Lines.Add($line)}
 }
 if($s.Lines.Count -eq 0){throw 'За выбранный месяц в колонке 22 не найдены записи. Проверьте вкладку и месяц.'}
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
 if($final){foreach($l in $s.Lines){if(-not $l.Reviewed){$errors.Add("Не проверен договор $($l.Contract).")}elseif($null -eq $s.Reference -and -not (Text $l.Name)){$errors.Add('Есть позиция без наименования.')}};if($s.Warnings.Count -gt 0 -and -not $s.Confirmed){$errors.Add('Просмотрите записи с неопределённым периодом и поставьте флажок подтверждения.')}}
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
