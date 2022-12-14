$file = "\\sidomain.local\NETLOGON\ADP\ADP Workforce Now.url"
$dst = "\\si-fp01\users$\"
$folder = ""
foreach ($item in Get-ChildItem($dst) -Directory)
{
    $folder = ""
    $folder = $dst + $item + "\Desktop"
    Write-Output($folder)
}
  
