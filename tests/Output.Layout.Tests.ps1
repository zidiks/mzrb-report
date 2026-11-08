$ErrorActionPreference='Stop'
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'app\Core.ps1')
$count=0
function Assert($ok,$message){if(-not $ok){throw ('FAILED: '+$message)};$script:count++}
function Row($id,[bool]$main,[string]$parent=''){
 return [pscustomobject]@{Row=$id;Identity=('basis:'+$id);ParentRow=0;ParentKey=$parent;IsMain=$main;IsRedistribution=$false;Values=(New-Object 'object[]' 13)}
}
$rows=New-Object Collections.Generic.List[object]
$empty=Row 9 $true;$empty.Values[0]='1';$empty.Values[2]='Supplier';$empty.Values[3]='D-1';$empty.Values[4]='01.01.2026';$empty.Values[6]='"';$rows.Add($empty)
$drug=Row 10 $true;$drug.Values[0]='5';$drug.Values[1]='Drug A';$drug.Values[3]='D-2';$drug.Values[5]=100;$rows.Add($drug)
$institution=Row 11 $false $drug.Identity;$institution.Values[6]='Institution';$institution.Values[7]=20;$rows.Add($institution)
$parent=Row 12 $true;$parent.Values[0]='10';$parent.Values[2]='Supplier';$parent.Values[3]='D-3';$rows.Add($parent)
$child=Row 13 $false $parent.Identity;$child.Values[6]='Second institution';$child.Values[7]=5;$rows.Add($child)
$offset=Row 14 $true;$offset.Values[0]='15';$offset.Values[3]='D-4';$offset.Values[5]=10;$offset.Values[12]=-10;$rows.Add($offset)
$blank=Row 15 $false $offset.Identity;$rows.Add($blank)
$changes=New-Object Collections.Generic.List[object]
Add-OutputRedistribution $rows @($drug) $drug ([decimal]-3) $changes @(5)
$layout=Finalize-OutputRows $rows
Assert ($layout.Removed.Count -eq 2 -and @($layout.Removed | Where-Object {$_.BasisRow -eq 9}).Count -eq 1) 'Remove metadata-only and fully empty rows'
Assert ($layout.PositionCount -eq 3) 'Keep named, financially meaningful and institution-parent positions'
Assert ($layout.Rows[0].Values[0] -eq '1' -and $layout.Rows[3].Values[0] -eq '2' -and $layout.Rows[-1].Values[0] -eq '3') 'Number retained positions consecutively at the end'
Assert ($layout.Rows[1].IsRedistribution -and $layout.Rows[1].Values[5] -eq -3 -and -not $layout.Rows[1].Values[0]) 'Redistribution follows its parent without a number'
Assert ($layout.Rows[2].Values[6] -eq 'Institution' -and $layout.Rows[2].Values[7] -eq 20) 'Institution continuation survives'
Assert ($layout.Rows[3].Values[3] -eq 'D-3' -and $layout.Rows[4].Values[7] -eq 5) 'Preserve metadata parent needed for institution data'
Assert ($layout.Rows[-1].Values[5] -eq 10 -and $layout.Rows[-1].Values[12] -eq -10) 'A zero-net row with nonzero components must survive'
Assert ($drug.Values[5] -eq 100) 'Redistribution never changes main shipment'
# Update and relocate an existing redistribution instead of appending a duplicate.
Add-OutputRedistribution $rows @($drug) $drug ([decimal]-4) $changes @(5)
Assert (@($rows | Where-Object {$_.IsRedistribution}).Count -eq 1 -and $rows[2].Values[5] -eq -4) 'Reuse an existing redistribution row'
Add-OutputRedistribution $rows @($drug) $drug ([decimal]0) $changes @(5)
$layout=Finalize-OutputRows $rows
Assert (@($layout.Rows | Where-Object {$_.IsRedistribution}).Count -eq 0) 'Remove cleared redistribution rows'
Assert ($layout.Rows[1].Values[7] -eq 20) 'Removing zero redistribution does not remove institutions'
# Added contracts have no original row number but must participate in numbering.
$new=Row 0 $true;$new.Identity='new:contract';$new.Values[1]='Added drug';$new.Values[3]='NEW';$new.Values[5]=30;$rows.Add($new)
Add-OutputRedistribution $rows @($new) $new ([decimal]-2) $changes @(6)
$layout=Finalize-OutputRows $rows
Assert ($layout.Rows[-2].Values[0] -eq '4' -and -not $layout.Rows[-1].Values[0] -and $layout.Rows[-1].Values[5] -eq -2) 'Added contract and redistribution receive correct final layout'
$institutionSum=[decimal]0;foreach($r in $layout.Rows){$institutionSum+=Money $r.Values[7]}
Assert ($institutionSum -eq 25) 'Institution subtotal conserved across cleanup and insertion'
"$count output layout checks passed"
