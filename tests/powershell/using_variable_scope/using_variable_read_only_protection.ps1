# vybe-test: powershell/using_variable_scope/using_variable_read_only_protection
$readOnlyVal = "OriginalValue"
$sb = {
    try {
        $using:readOnlyVal = "Mutated"
        return "Allowed"
    } catch {
        return "Protected"
    }
}
$res = &$sb
if ($readOnlyVal -ne "OriginalValue") {
    Write-Host "FAIL: using variable mutated caller variable state"
    exit 1
}
Write-Host "PASS"
exit 0
