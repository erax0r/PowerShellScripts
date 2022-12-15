$fileName = "ADP Workforce Now.url"
$filePath = "\\{server}\c$\Temp\"
$dst = "\\{server}\d$\Users\"
$folder = ""
foreach ($item in Get-ChildItem($dst) -Directory)
{
    $folder = ""
    $folder = $dst + $item + "\Desktop\"
    #Copy-Item -Path $file -Destination $folder
    #Write-Output "Copy-Item -Path $filePath -Destination $folder $fileName"
    Write-Output "robocopy `"$filePath`" `"$folder`" `"$fileName`" /COPYALL"
    #robocopy `"$filePath`" `"$folder`" `"$fileName`" /COPYALL
}
