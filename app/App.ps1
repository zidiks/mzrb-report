param(
 [ValidateSet('UI','Preview','Build')][string]$Mode='UI',
 [string]$Basis,[string]$BasisSheet,[string]$Source,[string]$Sheet,[string]$Period,[string]$Output,[string]$PreviewPath,
 [ValidateSet('Input','Done','Review','Busy','Error')][string]$Screen='Input',
 [string]$UpdatedFrom
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Core.ps1')
. (Join-Path $PSScriptRoot 'State.ps1')
. (Join-Path $PSScriptRoot 'Update.ps1')
$script:Ru=[Globalization.CultureInfo]::GetCultureInfo('ru-RU')
[Threading.Thread]::CurrentThread.CurrentCulture=$script:Ru
[Threading.Thread]::CurrentThread.CurrentUICulture=$script:Ru
if($Mode -eq 'Build'){
 $date=if($Period){[datetime]::ParseExact($Period,'yyyy-MM',$script:Invariant)}else{Sheet-Period $Sheet}
 if(-not $Basis -or -not $BasisSheet){throw 'Укажите -Basis (отчёт 1С) и -BasisSheet (вкладка 1С).'}
 $s=Analyze-Inputs $Source $Sheet $date $Basis $BasisSheet $true
 if(-not $Output){$Output=Unique-Output $s (Needs-Review $s)}
 $path=Save-Report $s $Output (-not (Needs-Review $s))
 [pscustomobject]@{Output=$path;Reference=($null -ne $s.Reference);Groups=$s.Groups.Count;Lines=$s.Lines.Count;Total=$s.ResultTotal;NeedsReview=(Needs-Review $s)} | ConvertTo-Json
 exit 0
}
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Data
$script:BasisPath='';$script:SourcePath='';$script:Session=$null;$script:Job=$null;$script:ResultPath='';$script:PeriodDate=$null;$script:Loading=$false;$script:AutoOpen=$Mode -eq 'UI';$script:LastError='';$script:ReviewWindow=$null
function Load-Xaml([string]$file){$xml=[xml](Get-Content -LiteralPath (Join-Path $PSScriptRoot $file) -Raw -Encoding UTF8);$reader=New-Object Xml.XmlNodeReader $xml;try{return [Windows.Markup.XamlReader]::Load($reader)}finally{$reader.Close()}}
function Fit-Window($target,[double]$availableWidth=0,[double]$availableHeight=0){
 # WPF work-area dimensions are device-independent and already account for Windows scaling.
 if($availableWidth -le 0){$availableWidth=[Windows.SystemParameters]::WorkArea.Width}
 if($availableHeight -le 0){$availableHeight=[Windows.SystemParameters]::WorkArea.Height}
 $target.MinWidth=[Math]::Min($target.MinWidth,$availableWidth)
 $target.MinHeight=[Math]::Min($target.MinHeight,$availableHeight)
 $target.MaxWidth=$availableWidth;$target.MaxHeight=$availableHeight
 $target.Width=[Math]::Min($target.Width,$availableWidth)
 $target.Height=[Math]::Min($target.Height,$availableHeight)
}
function Money-Display($v){return (Money $v).ToString('N2',$script:Ru)}
function Month-Display([datetime]$date){$text=$date.ToString('MMMM yyyy',$script:Ru);return $script:Ru.TextInfo.ToTitleCase($text)}
function Set-Visible($control,[bool]$value){$control.Visibility=if($value){'Visible'}else{'Collapsed'}}
function Show-Message([string]$title,[string]$body){$MessageTitle.Text=$title;$MessageBody.Text=$body;Set-Visible $MessagePanel $true;Set-Visible $ResultPanel $false}
function Clear-Message {Set-Visible $MessagePanel $false;$script:LastError=''}
function Log-Error($record){try{[IO.File]::WriteAllText((Join-Path $PSScriptRoot 'Data\Последняя ошибка.txt'),$record.ToString()+"`r`n"+$record.ScriptStackTrace,[Text.Encoding]::UTF8)}catch{}}
function Show-Error($record){
 Log-Error $record;$error=$record.Exception;while($null -ne $error.InnerException){$error=$error.InnerException};$text=$error.Message;if(-not $text){$text=$record.ToString()}
 if($text -match '80040154|REGDB_E_CLASSNOTREG'){$text='Для формирования отчёта нужен установленный Microsoft Excel.'}
 $script:LastError=$text
 if($Sheets.Items.Count -eq 0){$ReportMonth.Text='—';$script:PeriodDate=$null;Show-Message 'Не удалось прочитать исходный файл' $text}else{Show-Message 'Не удалось сформировать отчёт' $text}
 $Generate.Content='Попробовать ещё раз';Update-Action
}
function Save-Settings {
 try{[ordered]@{Source=$script:SourcePath;Sheet=[string]$Sheets.SelectedItem;Basis=$script:BasisPath;BasisSheet=[string]$BasisSheets.SelectedItem} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'Data\Settings.json') -Encoding UTF8}catch{}
}
function Update-Action {
 $busy=$null -ne $script:Job;$ChooseFile.IsEnabled=-not $busy;$Sheets.IsEnabled=-not $busy -and $Sheets.Items.Count -gt 0;$ManualPeriod.IsEnabled=-not $busy;$ResetReport.IsEnabled=-not $busy
 $ChooseBasis.IsEnabled=-not $busy;$BasisSheets.IsEnabled=-not $busy -and $BasisSheets.Items.Count -gt 0
 $Generate.IsEnabled=$script:BasisPath -ne '' -and $null -ne $BasisSheets.SelectedItem -and -not $busy -and $script:SourcePath -ne '' -and $null -ne $script:PeriodDate -and $null -ne $Sheets.SelectedItem
 $CheckUpdates.IsEnabled=-not $busy
 if($busy){$Generate.Content='Формируем отчёт…'}elseif($script:ResultPath){$Generate.Content='Открыть отчёт'}elseif($null -ne $script:Session -and (Needs-Review $script:Session)){$Generate.Content='Продолжить проверку'}elseif($script:LastError){$Generate.Content='Попробовать ещё раз'}else{$Generate.Content='Сформировать отчёт'}
}
function Clear-Result {$script:ResultPath='';$script:Session=$null;Set-Visible $ResultPanel $false;Set-Visible $OpenFolder $false;Set-Visible $ResetReport $false;Clear-Message;$DestinationHint.Text='Сохранится в папку «Готовые отчёты» рядом с файлом расчётов.'}
function Update-Period {
 if($script:Loading){return};Clear-Result;$script:PeriodDate=$null
 if($null -ne $Sheets.SelectedItem){try{$script:PeriodDate=Sheet-Period ([string]$Sheets.SelectedItem);Set-Visible $ManualPeriodPanel $false}catch{Set-Visible $ManualPeriodPanel $true;try{$script:PeriodDate=[datetime]::ParseExact($ManualPeriod.Text,'MM.yyyy',$script:Invariant)}catch{}}}
 $ReportMonth.Text=if($null -ne $script:PeriodDate){Month-Display $script:PeriodDate}else{'Укажите месяц'};Save-Settings;Update-Action
}
function Start-Work([string]$operation,$payload,[string]$message,[scriptblock]$done,[scriptblock]$failure=$null){
 if($null -ne $script:Job){return};Clear-Message;Set-Visible $ProgressPanel $true;$ProgressText.Text=$message
 $runspace=[RunspaceFactory]::CreateRunspace();$runspace.ApartmentState='STA';$runspace.ThreadOptions='ReuseThread';$runspace.Open()
 $shell=[PowerShell]::Create();$shell.Runspace=$runspace
 $body={param($root,$operation,$payload)
  $ErrorActionPreference='Stop'
  . (Join-Path $root 'Core.ps1');. (Join-Path $root 'State.ps1');. (Join-Path $root 'Update.ps1')
  switch($operation){
   'Sheets' {[pscustomobject]@{Data=@(Read-SheetNames $payload.Source)}}
   'Analyze' {$s=Analyze-Inputs $payload.Source $payload.Sheet $payload.Period $payload.Basis $payload.BasisSheet;[pscustomobject]@{Data=$s}}
   'Save' {$path=Unique-Output $payload.Session $payload.Draft;$saved=Save-Report $payload.Session $path (-not $payload.Draft);Save-Cache $payload.Session;[pscustomobject]@{Data=$saved;Draft=$payload.Draft;Total=$payload.Session.ResultTotal;Positions=$payload.Session.ResultPositions}}
   'UpdateCheck' {[pscustomobject]@{Data=(Latest-Release $root)}}
   'UpdateDownload' {[pscustomobject]@{Data=(Download-Update $payload.Release)}}
  }
 }
 [void]$shell.AddScript($body.ToString()).AddArgument($PSScriptRoot).AddArgument($operation).AddArgument($payload)
 $handle=$shell.BeginInvoke();$script:Job=[pscustomobject]@{Shell=$shell;Runspace=$runspace;Handle=$handle;Done=$done;Failure=$failure};Update-Action;$script:Timer.Start()
}
function Complete-Work {
 if($null -eq $script:Job -or -not $script:Job.Handle.IsCompleted){return}
 $script:Timer.Stop();$job=$script:Job;$script:Job=$null
 try{$results=$job.Shell.EndInvoke($job.Handle);if($job.Shell.HadErrors){throw $job.Shell.Streams.Error[0]};if($results.Count -ne 1){throw 'Не удалось получить результат операции.'};$result=$results[0]}
 catch{$result=$null;if($null -ne $job.Failure){& $job.Failure $_}else{Show-Error $_}}
 finally{$job.Shell.Dispose();$job.Runspace.Dispose();Set-Visible $ProgressPanel $false;Update-Action}
 if($null -ne $result){try{& $job.Done $result}catch{Show-Error $_}}
}
function Check-AppUpdates {
 $UpdateStatus.Text='Проверяем обновления…'
 Start-Work 'UpdateCheck' @{} 'Проверяем обновления на GitHub…' {
  param($result)
  $release=$result.Data
  if(-not $release.Available){$UpdateStatus.Text='Установлена последняя версия';return}
  $script:PendingRelease=$release;$UpdateStatus.Text='Доступна версия '+$release.Version
  $answer=[Windows.MessageBox]::Show($window,('Доступна версия '+$release.Version+'. Скачать и установить? Программа перезапустится. Файлы отчётов и сохранённая проверка останутся на месте.'),'Обновление программы','YesNo','Question','No')
  if($answer -ne 'Yes'){return}
  if($null -ne $script:Session -and $script:Session.Source){Save-Cache $script:Session}
  Start-Work 'UpdateDownload' @{Release=$script:PendingRelease} 'Скачиваем и проверяем обновление…' {
   param($download)
   Start-AppUpdate $download.Data $PSScriptRoot $PID
   $window.Close()
  } {param($errorRecord);Log-Error $errorRecord;$UpdateStatus.Text='Обновление не установлено. Повторите попытку.'}
 } {param($errorRecord);Log-Error $errorRecord;$UpdateStatus.Text='Не удалось проверить. Проверьте подключение к интернету.'}
}
function Load-Source([string]$path,[string]$preferredSheet=''){
 if($null -ne $script:Job){return}
 if(-not (Test-Path -LiteralPath $path -PathType Leaf)){Show-Message 'Выберите исходный файл' 'Файл не найден. Нажмите «Выбрать файл» и укажите книгу Excel.';return}
 if([IO.Path]::GetExtension($path).ToLowerInvariant() -notin @('.xls','.xlsx','.xlsm')){Show-Message 'Нужна книга Excel' 'Выберите файл с расширением .xlsx, .xls или .xlsm.';return}
 Clear-Result;$script:SourcePath=[IO.Path]::GetFullPath($path);$FileName.Text=[IO.Path]::GetFileName($path);$FileName.ToolTip=$script:SourcePath;$FileLocation.Text=[IO.Path]::GetDirectoryName($path);$FileLocation.ToolTip=$script:SourcePath
 $script:Loading=$true;$Sheets.Items.Clear();$script:Loading=$false;$script:PeriodDate=$null;$ReportMonth.Text='Загружаем вкладки…';$script:PreferredSheet=$preferredSheet
 Start-Work 'Sheets' @{Source=$script:SourcePath} 'Читаем список вкладок…' {
  param($result)
  $script:Loading=$true;foreach($name in $result.Data){[void]$Sheets.Items.Add([string]$name)};$script:Loading=$false
  if($Sheets.Items.Contains($script:PreferredSheet)){$Sheets.SelectedItem=$script:PreferredSheet}elseif($Sheets.Items.Count){$Sheets.SelectedIndex=$Sheets.Items.Count-1}
  Update-Period
 }
}
function Choose-Source {
 $dialog=New-Object Microsoft.Win32.OpenFileDialog;$dialog.Title='Выберите исходный отчёт МЗ РБ';$dialog.Filter='Книги Excel|*.xlsx;*.xls;*.xlsm';$dialog.CheckFileExists=$true
 if($dialog.ShowDialog($window) -eq $true){Load-Source $dialog.FileName}
}
function Load-Basis([string]$path,[string]$preferred=''){
 if($null -ne $script:Job){return}
 if(-not (Test-Path -LiteralPath $path -PathType Leaf) -or [IO.Path]::GetExtension($path).ToLowerInvariant() -notin @('.xls','.xlsx','.xlsm')){Show-Message 'Нужен отчёт из 1С' 'Выберите сохранённую книгу Excel из 1С.';return}
 Clear-Result;$script:BasisPath=[IO.Path]::GetFullPath($path);$BasisName.Text=[IO.Path]::GetFileName($path);$BasisName.ToolTip=$script:BasisPath;$BasisSheets.Items.Clear();$script:PreferredBasisSheet=$preferred
 Start-Work 'Sheets' @{Source=$script:BasisPath} 'Читаем вкладки отчёта 1С…' {
  param($result)
  foreach($name in $result.Data){[void]$BasisSheets.Items.Add([string]$name)}
  if($BasisSheets.Items.Contains($script:PreferredBasisSheet)){$BasisSheets.SelectedItem=$script:PreferredBasisSheet}elseif($BasisSheets.Items.Count){$BasisSheets.SelectedIndex=0}
  Save-Settings;Update-Action
  if($script:InitialSource){$path=$script:InitialSource;$script:InitialSource='';Load-Source $path $script:InitialSheet}
 }
}
function Choose-Basis {
 $dialog=New-Object Microsoft.Win32.OpenFileDialog;$dialog.Title='Выберите отчёт, выгруженный из 1С';$dialog.Filter='Книги Excel|*.xls;*.xlsx;*.xlsm';$dialog.CheckFileExists=$true
 if($dialog.ShowDialog($window) -eq $true){Load-Basis $dialog.FileName}
}
function Show-Result([string]$path,[bool]$draft){
 $script:ResultPath=$path;Set-Visible $MessagePanel $false;Set-Visible $ResultPanel $true;Set-Visible $OpenFolder $true;Set-Visible $ResetReport $true
 $ResultTitle.Text=if($draft){'Черновик готов'}else{'Отчёт готов'}
 $ResultDetails.Text=(Month-Display $script:PeriodDate)+' · '+$(if($script:Session.ResultPositions){$script:Session.ResultPositions}else{$script:Session.Lines.Count})+' позиций'+$(if($draft){' · требуется проверка'}else{''})
 $ResultFile.Text=[IO.Path]::GetFileName($path);$ResultFile.ToolTip=$path;$ResultTotal.Text=(Money-Display $(if($null -ne $script:Session.ResultTotal){$script:Session.ResultTotal}else{Session-Total $script:Session}))+' руб.'
 $DestinationHint.Text=if($draft){'Черновик можно проверить и исправить в Excel.'}else{'Выберите другую вкладку, чтобы сформировать следующий отчёт.'}
 Update-Action
 if($script:AutoOpen){Start-Process -FilePath $path}
}
function Begin-Save([bool]$draft){Start-Work 'Save' @{Session=$script:Session;Draft=$draft} 'Сохраняем отчёт в Excel…' {param($result);$script:Session | Add-Member NoteProperty ResultTotal $result.Total -Force;$script:Session | Add-Member NoteProperty ResultPositions $result.Positions -Force;Show-Result ([string]$result.Data) ([bool]$result.Draft)}}
function Continue-Session {
 if(Needs-Review $script:Session){
  $decision=Show-Review
  if($decision -eq 'Final'){Begin-Save $false}elseif($decision -eq 'Draft'){Begin-Save $true}else{$count=@(Pending-Groups $script:Session).Count;Show-Message 'Проверка сохранена' "Нужно проверить договоры: $count. Нажмите «Продолжить проверку», когда будете готовы.";Update-Action}
 }else{Begin-Save $false}
}
function Generate-Report {
 if($script:ResultPath){try{Start-Process -FilePath $script:ResultPath}catch{Show-Error $_};return}
 if($null -eq $script:PeriodDate -or $null -eq $Sheets.SelectedItem -or -not $script:BasisPath -or $null -eq $BasisSheets.SelectedItem){return}
 $current=$null -ne $script:Session -and (Session-Current $script:Session $script:SourcePath ([string]$Sheets.SelectedItem) $script:PeriodDate.ToString('yyyy-MM') $script:BasisPath $script:BasisPath ([string]$BasisSheets.SelectedItem))
 if($current){Continue-Session;return}
 Start-Work 'Analyze' @{Source=$script:SourcePath;Sheet=[string]$Sheets.SelectedItem;Period=$script:PeriodDate;Basis=$script:BasisPath;BasisSheet=[string]$BasisSheets.SelectedItem} 'Читаем выбранную вкладку и собираем отчёт…' {param($result);$script:Session=$result.Data;Continue-Session}
}
function Review-Table($group){
 $table=New-Object Data.DataTable;$table.Locale=$script:Ru
 [void]$table.Columns.Add('_Index',[int]);[void]$table.Columns.Add('Name',[string]);if($script:Session.BasisPath){[void]$table.Columns.Add('BasisRow',[int])};foreach($f in $script:Fields){[void]$table.Columns.Add($f,[decimal])}
 for($i=0;$i -lt $script:Session.Lines.Count;$i++){$line=$script:Session.Lines[$i];if($line.Group -ne $group.Id){continue};$row=$table.NewRow();$row['_Index']=$i;$row['Name']=$line.Name;if($script:Session.BasisPath){$row['BasisRow']=$line.BasisRow};foreach($f in $script:Fields){$row[$f]=Money $line.$f};$table.Rows.Add($row)}
 return ,$table
}
function Commit-Review {
 if($null -eq $script:ReviewTable -or $script:ReviewPage -ge $script:ReviewGroups.Count){return}
 if(-not $ReviewGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true) -or -not $ReviewGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)){throw 'Проверьте суммы: введите обычное число, например 1234,56.'}
 $changed=$false
 foreach($row in $script:ReviewTable.Rows){$line=$script:Session.Lines[[int]$row['_Index']];if($line.Name -cne [string]$row['Name'] -or $line.ContractDate -cne $ContractDate.Text.Trim()){$changed=$true};$line.Name=[string]$row['Name'];if($script:Session.BasisPath){if($line.BasisRow -ne [int]$row['BasisRow']){$changed=$true};$line.BasisRow=[int]$row['BasisRow']};foreach($f in $script:Fields){$value=Money $row[$f];if((Money $line.$f) -ne $value){$changed=$true};$line.$f=$value};$line.ContractDate=$ContractDate.Text.Trim()}
 if($changed){$group=$script:ReviewGroups[$script:ReviewPage];foreach($line in $script:Session.Lines){if($line.Group -eq $group.Id){$line.Reviewed=$false}}}
 Save-Cache $script:Session
}
function Update-Balance {
 if($null -eq $script:ReviewTable -or $script:ReviewPage -ge $script:ReviewGroups.Count){return}
 $group=$script:ReviewGroups[$script:ReviewPage];$expected=[decimal]0;foreach($e in $script:Session.Events){if($e.BlockId -eq $group.Id){$expected+=$e.Amount}}
 $actual=[decimal]0;foreach($row in $script:ReviewTable.Rows){foreach($f in $script:Fields){$actual+=Money $row[$f]}}
 $Balance.Text='По исходнику: '+(Money-Display $expected)+'   ·   В строках: '+(Money-Display $actual)+'   ·   Разница: '+(Money-Display ($actual-$expected))
 $Balance.Foreground=if($actual -eq $expected){'#167367'}else{'#A33C2B'}
}
function Show-ReviewPage {
 $ReviewError.Text='';$ReviewBack.IsEnabled=$script:ReviewPage -gt 0;$warningsPage=$script:ReviewPage -ge $script:ReviewGroups.Count
 $pageCount=$script:ReviewGroups.Count+$(if($script:HasWarningPage){1}else{0});$ReviewProgress.Text="Шаг $($script:ReviewPage+1) из $pageCount"
 Set-Visible $WarningPage $warningsPage;Set-Visible $ReviewGrid (-not $warningsPage);Set-Visible $TableHeading (-not $warningsPage);Set-Visible $ContractDatePanel (-not $warningsPage);Set-Visible $SourceExpander (-not $warningsPage);Set-Visible $Balance (-not $warningsPage)
 if($warningsPage){$ContractTitle.Text='Записи без отчётного месяца';$ContractHelp.Text='Подтвердите, что эти данные можно оставить за пределами отчёта.';$UndatedRows.Text=$script:Session.Warnings -join "`r`n`r`n";$WarningsConfirmed.IsChecked=$script:Session.Confirmed;$ReviewNext.Content='Сформировать отчёт';$script:ReviewTable=$null;return}
 $group=$script:ReviewGroups[$script:ReviewPage];$script:ReviewTable=Review-Table $group;$ReviewGrid.ItemsSource=$script:ReviewTable.DefaultView
 $ContractTitle.Text='Договор '+$group.Contract;$ContractHelp.Text=if($script:ReviewTable.Rows.Count -gt 1){'В договоре несколько препаратов. Проверьте, как общая сумма распределена между строками.'}else{'Проверьте реквизиты договора и распределение расходов.'}
 $first=$script:Session.Lines[[int]$script:ReviewTable.Rows[0]['_Index']];$ContractDate.Text=$first.ContractDate;$Explanation.Text=$group.Note;if($script:Session.BasisPath){$ContractHelp.Text+=' Строка 1С: номер строки исходной книги; 0 — не переносить сумму. Реквизиты существующих строк и данные учреждений сохраняются.'}
 $sourceTable=New-Object Data.DataTable;foreach($name in @('Строка','Запись','Сумма')){[void]$sourceTable.Columns.Add($name,[string])}
 foreach($ev in $script:Session.Events){if($ev.BlockId -eq $group.Id){$r=$sourceTable.NewRow();$r['Строка']=[string]$ev.Row;$r['Запись']=$ev.Label;$r['Сумма']=Money-Display $ev.Amount;$sourceTable.Rows.Add($r)}}
 foreach($ex in $script:Session.Expenses){if($ex.BlockId -eq $group.Id){foreach($field in @('Conversion18','Commission20','Vat19','Other21')){$names=@{Conversion18='Покупка валюты';Commission20='Комиссия';Vat19='НДС / таможня';Other21='Прочие расходы'};if((Money $ex.$field) -ne 0){$r=$sourceTable.NewRow();$r['Строка']=[string]$ex.Row;$r['Запись']='Расшифровка: '+$names[$field];$r['Сумма']=Money-Display $ex.$field;$sourceTable.Rows.Add($r)}}}}
 $SourceRows.ItemsSource=$sourceTable.DefaultView;$ReviewNext.Content=if($script:ReviewPage -eq $pageCount-1){'Подтвердить и сформировать'}else{'Подтвердить и дальше'}
 $script:ReviewTable.Add_ColumnChanged({Update-Balance});Update-Balance
}
function Confirm-ReviewPage {
 try{
  if($script:ReviewPage -ge $script:ReviewGroups.Count){if($WarningsConfirmed.IsChecked -ne $true){throw 'Поставьте отметку, что записи просмотрены.'};$script:Session.Confirmed=$true}else{
   Commit-Review;$group=$script:ReviewGroups[$script:ReviewPage];$date=Parse-Date $ContractDate.Text;if($null -eq $date){throw 'Укажите дату договора в формате дд.мм.гггг.'}
   $lines=@($script:Session.Lines | Where-Object {$_.Group -eq $group.Id});foreach($line in $lines){if(-not (Text $line.Name)){throw 'Укажите наименование каждого препарата.'};$line.ContractDate=$date.ToString('dd.MM.yyyy')}
   $errors=@(Validate-Session $script:Session $false | Where-Object {$_.StartsWith('Договор '+$group.Contract+':')});if($errors.Count){throw ($errors -join "`r`n")}
   foreach($line in $lines){$line.Reviewed=$true}
  }
  Save-Cache $script:Session;$script:ReviewPage++
  $total=$script:ReviewGroups.Count+$(if($script:HasWarningPage){1}else{0})
  if($script:ReviewPage -ge $total){$script:Decision='Final';$script:ReviewWindow.Close()}else{Show-ReviewPage}
 }catch{$ReviewError.Text=$_.Exception.Message}
}
function Show-Review {
 $script:Decision='Cancel';$script:ReviewGroups=@(Pending-Groups $script:Session);$script:HasWarningPage=$script:Session.Warnings.Count -gt 0 -and -not $script:Session.Confirmed;$script:ReviewPage=0
 $script:ReviewWindow=Load-Xaml 'Review.xaml';Fit-Window $script:ReviewWindow;$script:ReviewWindow.Owner=$window;$script:ReviewWindow.Icon=$window.Icon;[void]$script:ReviewWindow.Resources.MergedDictionaries.Add($window.Resources)
 foreach($name in @('ReviewProgress','ContractTitle','ContractHelp','ContractDatePanel','ContractDate','TableHeading','ReviewGrid','WarningPage','UndatedRows','WarningsConfirmed','Balance','ReviewError','SourceExpander','Explanation','SourceRows','ReviewBack','ReviewCancel','SaveDraft','ReviewNext','AddDrug')){Set-Variable -Scope Script -Name $name -Value $script:ReviewWindow.FindName($name)}
 $titles=@{Name='Лекарственное средство';Transfer='Отпуск';Redistribution='Перераспр.';Conversion='Покупка валюты';Fx='Курс. разницы';Commission='Комиссия';Discount='Скидки';Other='Прочие'}
 foreach($field in @($(if($script:Session.BasisPath){'BasisRow'}),'Name')+$script:Fields | Where-Object {$_}){$col=New-Object Windows.Controls.DataGridTextColumn;$col.Header=if($field -eq 'BasisRow'){'Строка 1С (0 — пропуск)'}else{$titles[$field]};$col.Width=if($field -eq 'Name'){210}else{100};$binding=New-Object Windows.Data.Binding($field);$binding.UpdateSourceTrigger='PropertyChanged';$binding.ValidatesOnExceptions=$true;$binding.ConverterCulture=$script:Ru;if($field -notin @('Name','BasisRow')){$binding.StringFormat='N2'};$col.Binding=$binding;[void]$ReviewGrid.Columns.Add($col)}
 $ReviewNext.Add_Click({Confirm-ReviewPage});$ReviewCancel.Add_Click({try{Commit-Review}catch{};$script:ReviewWindow.Close()})
 $ReviewBack.Add_Click({try{Commit-Review;if($script:ReviewPage -gt 0){$script:ReviewPage--;Show-ReviewPage}}catch{$ReviewError.Text=$_.Exception.Message}})
 $SaveDraft.Add_Click({try{Commit-Review;$errors=@(Validate-Session $script:Session $false);if($errors.Count){throw 'Суммы позиций пока не совпадают с исходником. Исправьте распределение, чтобы сохранить черновик.'};Save-Cache $script:Session;$script:Decision='Draft';$script:ReviewWindow.Close()}catch{$ReviewError.Text=$_.Exception.Message}})
 $AddDrug.Add_Click({try{Commit-Review;$group=$script:ReviewGroups[$script:ReviewPage];$existing=@($script:Session.Lines | Where-Object {$_.Group -eq $group.Id});if(@($existing | Where-Object {-not (Text $_.Name)}).Count){return};$first=$existing[0];$props=[ordered]@{};foreach($prop in $first.PSObject.Properties){$props[$prop.Name]=$prop.Value};$line=[pscustomobject]$props;$line.Name='';$line.Reviewed=$false;if($script:Session.BasisPath){$line.BasisRow=-1};$line | Add-Member -NotePropertyName NewlyAdded -NotePropertyValue $true -Force;foreach($f in $script:Fields){$line.$f=[decimal]0};$index=0;for($i=0;$i -lt $script:Session.Lines.Count;$i++){if($script:Session.Lines[$i].Group -eq $group.Id){$index=$i+1}};$script:Session.Lines.Insert($index,$line);Show-ReviewPage}catch{$ReviewError.Text=$_.Exception.Message}})
 $script:ReviewWindow.Add_Closing({try{Commit-Review;foreach($line in @($script:Session.Lines | Where-Object {$_.NewlyAdded -and -not (Text $_.Name) -and (Line-Total $_) -eq 0})){[void]$script:Session.Lines.Remove($line)};Save-Cache $script:Session}catch{}})
 Show-ReviewPage
 if($Mode -eq 'Preview' -and $Screen -eq 'Review'){Capture-Window $script:ReviewWindow $PreviewPath;$script:ReviewWindow=$null;return 'Cancel'}
 [void]$script:ReviewWindow.ShowDialog();$script:ReviewWindow=$null;return $script:Decision
}
function Capture-Window($target,[string]$path){
 $target.WindowStartupLocation='Manual';$target.Left=-20000;$target.Top=-20000;$target.Show();$target.UpdateLayout()
 $visual=$target.Content;$width=[int]($visual.ActualWidth+$visual.Margin.Left+$visual.Margin.Right);$height=[int]($visual.ActualHeight+$visual.Margin.Top+$visual.Margin.Bottom)
 $bmp=New-Object Windows.Media.Imaging.RenderTargetBitmap($width,$height,96,96,[Windows.Media.PixelFormats]::Pbgra32);$bmp.Render($target)
 $encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder;$encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bmp));$stream=[IO.File]::Create($path);try{$encoder.Save($stream)}finally{$stream.Dispose()};$target.Close()
}
function Pump-Work([int]$seconds=55){
 $deadline=[datetime]::UtcNow.AddSeconds($seconds)
 while($null -ne $script:Job -and [datetime]::UtcNow -lt $deadline){[Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background);Start-Sleep -Milliseconds 30}
 if($null -ne $script:Job){throw 'Timed out waiting for background work'}
 if($script:LastError){throw $script:LastError}
}
try{
 $window=Load-Xaml 'Main.xaml';Fit-Window $window;$icon=Join-Path $PSScriptRoot 'Data\Report.ico';if(Test-Path -LiteralPath $icon){$window.Icon=[Windows.Media.Imaging.BitmapFrame]::Create([uri]$icon)}
 foreach($name in @('BasisName','BasisSheets','ChooseBasis','FileCard','FileName','FileLocation','ChooseFile','Sheets','ReportMonth','ManualPeriodPanel','ManualPeriod','ProgressPanel','ProgressText','MessagePanel','MessageTitle','MessageBody','ResultPanel','ResultTitle','ResultDetails','ResultFile','ResultTotal','Generate','DestinationHint','OpenFolder','ResetReport','CheckUpdates','VersionLabel','UpdateStatus')){Set-Variable -Scope Script -Name $name -Value $window.FindName($name)}
 $VersionLabel.Text='Версия '+(App-Version).ToString();$CheckUpdates.Add_Click({try{Check-AppUpdates}catch{Log-Error $_;$UpdateStatus.Text='Не удалось установить обновление. Повторите попытку.'}})
 if($UpdatedFrom -eq 'GitHub'){$UpdateStatus.Text='Обновление установлено'}
 $ChooseBasis.Add_Click({Choose-Basis});$BasisSheets.Add_SelectionChanged({Clear-Result;Save-Settings;Update-Action})
 $Generate.IsDefault=$true;$ChooseFile.Add_Click({Choose-Source});$Generate.Add_Click({try{Generate-Report}catch{Show-Error $_}});$Sheets.Add_SelectionChanged({Update-Period});$ManualPeriod.Add_TextChanged({Update-Period})
 $OpenFolder.Add_Click({if($script:ResultPath){Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"'+$script:ResultPath+'"')}})
 $ResetReport.Add_Click({try{$script:ResultPath='';Set-Visible $ResultPanel $false;Set-Visible $OpenFolder $false;Set-Visible $ResetReport $false;Generate-Report}catch{Show-Error $_}})
 $FileCard.Add_PreviewDragOver({param($sender,$e);$e.Effects=if($e.Data.GetDataPresent([Windows.DataFormats]::FileDrop) -and $null -eq $script:Job){'Copy'}else{'None'};$e.Handled=$true})
 $FileCard.Add_Drop({param($sender,$e);if($null -eq $script:Job -and $e.Data.GetDataPresent([Windows.DataFormats]::FileDrop)){$files=$e.Data.GetData([Windows.DataFormats]::FileDrop);if($files.Count -eq 1){Load-Source $files[0]}else{Show-Message 'Нужен один исходный файл' 'Перетащите одну книгу Excel.'}}})
 $script:Timer=New-Object Windows.Threading.DispatcherTimer;$script:Timer.Interval=[timespan]::FromMilliseconds(120);$script:Timer.Add_Tick({Complete-Work})
 $window.Add_Closing({param($sender,$e);if($null -ne $script:Job){$e.Cancel=$true;$ProgressText.Text='Дождитесь завершения операции с Excel…'}elseif($null -ne $script:Session -and $script:Session.Source){Save-Cache $script:Session}})
 if($Mode -eq 'Preview'){
  $script:SourcePath=if($Source){$Source}else{'C:\Отчёты\Отчёт МЗ РБ октябрь 2026.xlsx'};$FileName.Text=[IO.Path]::GetFileName($script:SourcePath);$FileLocation.Text=[IO.Path]::GetDirectoryName($script:SourcePath)
  $script:Loading=$true;[void]$Sheets.Items.Add('01.10.26');[void]$Sheets.Items.Add('01.11.26');$Sheets.SelectedItem='01.11.26';$script:Loading=$false;Update-Period
  if($Screen -eq 'Done'){$script:Session=[pscustomobject]@{Lines=@(1..3 | ForEach-Object {[pscustomobject]@{Transfer=[decimal]0;Redistribution=[decimal]0;Conversion=[decimal]0;Fx=[decimal]0;Commission=[decimal]0;Discount=[decimal]0;Other=[decimal]0}})};$script:Session.Lines[0].Transfer=[decimal]12345.67;Show-Result 'C:\Отчёты\Готовые отчёты\102026_отчёт.xls' $false}
  if($Screen -eq 'Busy'){Set-Visible $ProgressPanel $true;$ProgressText.Text='Собираем данные за октябрь…';$Generate.IsEnabled=$false;$Generate.Content='Формируем отчёт…'}
  if($Screen -eq 'Error'){Show-Message 'Не удалось прочитать исходный файл' 'Проверьте, что файл сохранён и доступен, затем выберите его ещё раз.'}
  if($Screen -eq 'Review'){$window.WindowStartupLocation='Manual';$window.Left=-20000;$window.Top=-20000;$window.Show();$script:Session=Analyze-BuiltIn $Source '01.10.26' ([datetime]'2026-09-01') $true;$null=Show-Review;$window.Close();exit 0}
  Capture-Window $window $PreviewPath;exit 0
 }
 $settingsPath=Join-Path $PSScriptRoot 'Data\Settings.json';$preferred=''
 if(Test-Path -LiteralPath $settingsPath){try{$settings=Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 | ConvertFrom-Json;if(Test-Path -LiteralPath $settings.Source -PathType Leaf){$Source=$settings.Source;$preferred=$settings.Sheet}}catch{}}
 if(-not $Source){foreach($folder in @($PSScriptRoot,(Split-Path $PSScriptRoot -Parent))){$candidate=Get-ChildItem -LiteralPath $folder -File -Filter '*.xlsx' | Where-Object {$_.Name -like '*Отч*' -and $_.Name -notlike '~$*'} | Sort-Object LastWriteTime -Descending | Select-Object -First 1;if($candidate){$Source=$candidate.FullName;break}}}
 if(-not $Basis -and $settings -and $settings.Basis){$Basis=$settings.Basis;$BasisSheet=$settings.BasisSheet}
 $script:InitialBasis=$Basis;$script:InitialBasisSheet=$BasisSheet
 $script:InitialSource=$Source;$script:InitialSheet=$preferred;$window.Add_ContentRendered({if($script:InitialBasis){$path=$script:InitialBasis;$script:InitialBasis='';Load-Basis $path $script:InitialBasisSheet}elseif($script:InitialSource){$path=$script:InitialSource;$script:InitialSource='';Load-Source $path $script:InitialSheet}})
 [void]$window.ShowDialog()
}catch{Log-Error $_;if($Mode -eq 'UI'){[void][Windows.MessageBox]::Show($_.Exception.Message,'Не удалось открыть программу','OK','Error')}else{Write-Output $_.ToString();Write-Output $_.ScriptStackTrace};exit 1}
finally{if($null -ne $script:Timer){$script:Timer.Stop()};if($null -ne $script:Job -and $script:Job.Handle.IsCompleted){$script:Job.Shell.Dispose();$script:Job.Runspace.Dispose()}}
