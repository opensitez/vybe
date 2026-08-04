# vybe-test: powershell/scriptblocks/foreach_object_scriptblock
$numbers = 1..5
$doubled = $numbers | ForEach-Object { $_ * 2 }
$expected = @(2, 4, 6, 8, 10)
for ($i = 0; $i -lt 5; $i++) {
    if ($doubled[$i] -ne $expected[$i]) {
        Write-Host "FAIL: index $i expected $($expected[$i]) got $($doubled[$i])"
        exit 1
    }
}
Write-Host "PASS"
exit 0
