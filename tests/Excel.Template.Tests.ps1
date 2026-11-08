$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'app\Core.ps1')
# GitHub-hosted Windows does not include Excel. Report this separately from a passing COM check.
if($null -eq [Type]::GetTypeFromProgID('Excel.Application')){
 Write-Output 'SKIPPED: real Excel template smoke test (Microsoft Excel is not installed)'
 return
}
$excel=$null;$book=$null
try{
 $excel=New-Excel
 # Strict normal open: a repaired open must not hide a regression in the shipped template.
 $book=$excel.Workbooks.Open((Join-Path $root 'app\Data\Layout.xls'),0,$true)
 $matrix=Read-Matrix $book.Worksheets.Item(1)
 if(-not (Text (Cell $matrix 6 3)).Contains('лекарственного')){throw 'Missing report heading'}
 if((Text (Cell $matrix 7 11)) -ne 'номер'){throw 'Missing contract heading'}
 if((Text (Cell $matrix 204 14)).Replace(' ','') -ne 'ВСЕГО:'){throw 'Missing report total'}
 Write-Output ('Excel '+$excel.Version+': bundled XLS opened without repair; headings and total checked')
}finally{
 if($null -ne $book){$book.Close($false)}
 Close-Excel $excel
}
