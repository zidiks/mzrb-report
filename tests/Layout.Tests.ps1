$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'app\Core.ps1')
# Managed test doubles cannot accept the COM Missing sentinel. Substitute only that sentinel.
$definition=(Get-Command Open-Layout).Definition.Replace('[Type]::Missing','$null')
Invoke-Expression ('function Open-Layout {'+$definition+'}')
$script:LayoutChecks=0
function Assert($ok,$message){if(-not $ok){throw $message};$script:LayoutChecks++}
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
public class LayoutTestBooks {
 public List<object[]> Calls = new List<object[]>();
 public bool FailNormal, FailRepair;
 public string Open(string path, int links, bool readOnly) {
  Calls.Add(new object[] {path, links, readOnly});
  if (FailNormal) throw new Exception("Office file validation rejected layout");
  return "opened workbook";
 }
 public string Open(string path, int links, bool readOnly, object d=null, object e=null, object f=null, object g=null, object h=null, object i=null, object j=null, object k=null, object l=null, object m=null, object n=null, object repair=null) {
  Calls.Add(new object[] {path, links, readOnly,d,e,f,g,h,i,j,k,l,m,n,repair});
  if (FailRepair) throw new Exception("Repair failed");
  return "opened workbook";
 }
}
'@
$books=New-Object LayoutTestBooks
$excel=[pscustomobject]@{Workbooks=$books}
Assert ((Open-Layout $excel 'layout.xls' $true) -eq 'opened workbook') 'Normal layout open failed'
Assert ($books.Calls.Count -eq 1) 'Healthy layout should not be repaired'
$books.Calls.Clear();$books.FailNormal=$true
Assert ((Open-Layout $excel 'layout.xls' $true) -eq 'opened workbook') 'Repair fallback failed'
Assert ($books.Calls.Count -eq 2 -and $books.Calls[1].Count -eq 15) 'Expected repair after normal failure'
Assert ($books.Calls[1][0] -eq 'layout.xls' -and $books.Calls[1][2] -eq $true -and $books.Calls[1][14] -eq 1) 'Repair must use xlRepairFile and preserve read-only mode'
$books.Calls.Clear()
[void](Open-Layout $excel 'output-copy.xls' $false)
Assert ($books.Calls[1][2] -eq $false) 'Output copy must be editable'
$books.FailRepair=$true;$failed=$false
try{Open-Layout $excel 'broken.xls' $true}catch{$failed=$_.Exception.Message.Contains('broken.xls')}
Assert $failed 'Repair failure must identify the template'
"$script:LayoutChecks layout checks passed"
