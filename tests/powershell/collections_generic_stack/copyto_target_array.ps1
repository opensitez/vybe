# vybe-test: powershell/collections_generic_stack/copyto_target_array
$s = [System.Collections.Generic.Stack[string]]::new()
$s.Push("bottom"); $s.Push("top")
[string[]]$target = New-Object string[] 2
$s.CopyTo($target, 0)
if ($target[0] -ne "top" -or $target[1] -ne "bottom") {
    Write-Host "FAIL: CopyTo on Stack failed"
    exit 1
}
Write-Host "PASS"
exit 0
