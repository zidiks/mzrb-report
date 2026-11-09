$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'app\Core.ps1')
. (Join-Path $root 'app\State.ps1')
if($null -eq [Type]::GetTypeFromProgID('Excel.Application')){Write-Output 'SKIPPED: two-file export round trip (Microsoft Excel is not installed)';return}
$dir=Join-Path ([IO.Path]::GetTempPath()) ('MzrbMerge-'+[guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($dir)
$excel=$null;$book=$null;$s=$null
function Assert($condition,$message){if(-not $condition){throw $message}}
try{
 $basisPath=Join-Path $dir 'basis.xls';$sourcePath=Join-Path $dir 'ledger.xlsx';$output=Join-Path $dir 'result.xls'
 Copy-Item -LiteralPath (Join-Path $root 'app\Data\Layout.xls') -Destination $basisPath
 $excel=New-Excel;$book=$excel.Workbooks.Open($basisPath,0,$false);$ws=$book.Worksheets.Item(1);$ws.Name='Report'
 [void]$ws.Range('A9:AH203').ClearContents();[void]$ws.Rows.Item('12:202').Delete()
 Set-ExcelText $ws 4 17 'за период с 01.10.2026 по 31.10.2026'
 foreach($row in @(9,11)){
  Set-ExcelText $ws $row 2 $(if($row -eq 9){'1'}else{'2'});Set-ExcelText $ws $row 3 $(if($row -eq 9){'Drug A'}else{'Drug B'})
  Set-ExcelText $ws $row 7 'Supplier';Set-ExcelText $ws $row 11 $(if($row -eq 9){'A-1'}else{'A-2'});Set-ExcelText $ws $row 14 '01.01.2026'
 }
 $ws.Cells.Item(9,16).Formula='1000';$ws.Cells.Item(9,28).Formula='50';$ws.Cells.Item(9,33).Formula='3';$ws.Cells.Item(11,16).Formula='200';$ws.Cells.Item(11,25).Formula='-2'
 Set-ExcelText $ws 9 18 'Institution A';$ws.Cells.Item(9,21).Formula='250';Set-ExcelText $ws 10 18 'Institution B';$ws.Cells.Item(10,21).Formula='100'
 Set-ExcelText $ws 13 14 'ВСЕГО:';$book.Save();$book.Close($false);$book=$null
 $book=$excel.Workbooks.Add();$ws=$book.Worksheets.Item(1);$ws.Name='01.11.26'
 Set-ExcelText $ws 1 22 'Отпущено ЛПУ';$ws.Cells.Item(5,1).Formula='1';Set-ExcelText $ws 5 7 'A-1';Set-ExcelText $ws 5 8 'Supplier';Set-ExcelText $ws 5 9 'Drug A'
 Set-ExcelText $ws 5 22 'Октябрь';$ws.Cells.Item(5,23).Formula='900';Set-ExcelText $ws 6 7 '01.01.2026';Set-ExcelText $ws 6 22 'ком.возн.10';$ws.Cells.Item(6,23).Formula='0';Set-ExcelText $ws 7 1 'Итого'
 $book.SaveAs($sourcePath,51);$book.Close($false);$book=$null;Close-Excel $excel;$excel=$null
 $beforeBasis=File-Sha256 $basisPath;$beforeSource=File-Sha256 $sourcePath
 $s=Analyze-Inputs $sourcePath '01.11.26' ([datetime]'2026-10-01') $basisPath 'Report' $true
 Assert ($s.Automatic -and -not (Needs-Review $s)) 'Native analysis must select automatic contract export without a wizard'
 [void](Save-Report $s $output $true)
 Assert ((File-Sha256 $basisPath) -eq $beforeBasis -and (File-Sha256 $sourcePath) -eq $beforeSource) 'Input workbook changed'
 $excel=New-Excel;$book=$excel.Workbooks.Open($output,0,$true);$ws=$book.Worksheets.Item(1)
 Assert ($book.Worksheets.Count -eq 1 -and $book.FileFormat -eq 56) 'Output must contain selected report as real XLS'
 Assert ($ws.Range('F9').Value2 -eq 900 -and $ws.Range('K9').Value2 -eq 0 -and $ws.Range('M9').Value2 -eq 0) 'Override/zero/preservation failed in output'
 Assert ($ws.Range('G9').Value2 -eq 'Institution A' -and $ws.Range('H9').Value2 -eq 250 -and $ws.Range('G10').Value2 -eq 'Institution B' -and $ws.Range('H10').Value2 -eq 100) 'Institution data lost'
 Assert ($ws.Range('B11').Value2 -eq 'Drug B' -and $ws.Range('F11').Value2 -eq 200) 'Unmatched base position lost'
 Assert ($ws.Range('F13').Value2 -eq 1098 -and $ws.Range('H12').Value2 -eq 350) 'Incorrect calculated totals'
 Assert ($ws.Range('F12').HasFormula -and $ws.Range('F13').HasFormula -and $ws.Range('H12').HasFormula) 'Totals must be formulas'
 Assert ($ws.Range('J6').Value2 -eq 'Курсовые разницы' -and $ws.Range('K6').Value2 -eq 'Комиссионное вознаграждение РУП БЕЛФАРМАЦИЯ') 'Headers not corrected'
 Assert ($ws.Range('F9').NumberFormat.Contains('#,##0.00')) 'Money number format missing'
 # Force invariant-format rejection while forwarding the localized setter to real Excel.
 foreach($address in @('F9','J11','K9')){
  $actual=$ws.Range($address);$before=$actual.Value2
  $probe=[pscustomobject]@{Application=$excel;Actual=$actual}
  $probe | Add-Member ScriptProperty NumberFormat {''} {param($value);throw 'Force localized-format path'}
  $probe | Add-Member ScriptProperty NumberFormatLocal {$this.Actual.NumberFormatLocal} {param($value);$this.Actual.NumberFormatLocal=$value}
  Set-ExcelMoneyFormat $probe $address
  Assert ($actual.Value2 -eq $before -and $actual.NumberFormat.Contains('#,##0.00')) 'Localized formatting must preserve positive, negative and zero values in real Excel'
 }
 Assert ($ws.Range('F13').MergeArea.Columns.Count -eq 2 -and $ws.Range('F13').MergeArea.NumberFormat.Contains('#,##0.00')) 'Inherited F:G grand total must retain its merge and receive the money format'
 Assert ($ws.Range('D9').MergeArea.Rows.Count -eq 2 -and $ws.Range('B9').MergeArea.Rows.Count -eq 2) 'Contract/product metadata not merged'
 Assert ($ws.Range('F17').HasFormula -and $ws.Range('F17').Value2 -eq 0 -and $ws.Range('F19').HasFormula -and $ws.Range('F19').Value2 -eq 1098) 'Redistribution footer formulas failed'
 Assert ($ws.UsedRange.Columns.Count -eq 13) 'Phantom columns remain'
 Assert (Test-Path -LiteralPath (Join-Path $dir 'Проверки\result.html')) 'Merge audit missing'
 Write-Output ('Excel '+$excel.Version+': two-file export round trip passed')
}finally{
 if($null -ne $book){$book.Close($false)};Close-Excel $excel
 # This test owns this GUID-named temporary directory, including the cached review file.
 if($s){$cache=Cache-Path $s.Source $s.Sheet $s.PeriodKey $s.BasisPath $s.BasisSheet;if(Test-Path -LiteralPath $cache){Remove-Item -LiteralPath $cache}}
 if(Test-Path -LiteralPath $dir){Remove-Item -LiteralPath $dir -Recurse -Force}
}
