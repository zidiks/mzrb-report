$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework
$root=Split-Path $PSScriptRoot -Parent
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'app\App.ps1'),[ref]$null,[ref]$null)
$fn=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Fit-Window'},$true)
Invoke-Expression $fn.Extent.Text
$count=0
foreach($size in @(@(1024,728),@(800,560),@(640,440))){
 foreach($name in @('Main.xaml','Review.xaml')){
  $reader=New-Object Xml.XmlNodeReader ([xml][IO.File]::ReadAllText((Join-Path $root ('app\'+$name))))
  try{$w=[Windows.Markup.XamlReader]::Load($reader)}finally{$reader.Close()}
  try{
   Fit-Window $w $size[0] $size[1]
   if($w.Width -gt $size[0] -or $w.Height -gt $size[1] -or $w.MinHeight -gt $size[1] -or $w.MinWidth -gt $size[0]){throw ('Window exceeds work area: '+$name)}
   if($w.Content -isnot [Windows.Controls.ScrollViewer]){throw ('Window must support scrolling: '+$name)}
   $w.Show();$w.UpdateLayout()
   if($name -eq 'Main.xaml'){$w.FindName('MessagePanel').Visibility='Visible';$w.FindName('MessageBody').Text=('Long error message '*50)}
   $w.UpdateLayout();$w.Content.ScrollToEnd();$w.UpdateLayout()
   $button=if($name -eq 'Main.xaml'){$w.FindName('CheckUpdates')}else{$w.FindName('ReviewNext')}
   $p=$button.TransformToAncestor($w.Content).Transform((New-Object Windows.Point(0,0)))
   if($p.Y -lt -1 -or $p.Y+$button.ActualHeight -gt $w.Content.ActualHeight+1){throw ('Bottom button inaccessible: '+$name)}
   $count++
  }finally{$w.Close()}
 }
}
"$count window sizes checked"
