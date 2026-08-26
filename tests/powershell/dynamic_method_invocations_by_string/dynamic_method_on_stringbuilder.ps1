# vybe-test: powershell/dynamic_method_invocations_by_string/dynamic_method_on_stringbuilder
$sb = [System.Text.StringBuilder]::new()
$m = "Append"
$null = $sb.$m("part1")
$null = $sb.$m("part2")
if ($sb.ToString() -ne "part1part2") {
    Write-Host "FAIL: Dynamic method on StringBuilder failed"
    exit 1
}
Write-Host "PASS"
exit 0
