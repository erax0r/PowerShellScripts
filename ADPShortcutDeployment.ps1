$file = "\\{domain}\NETLOGON\ADP\ADP Workforce Now.url"
$dst = "\\{server}\users$\"
$folder = ""
foreach ($item in Get-ChildItem($dst) -Directory)
{
    $folder = ""
    $folder = $dst + $item + "\Desktop"
    Write-Output($folder)
}
  
