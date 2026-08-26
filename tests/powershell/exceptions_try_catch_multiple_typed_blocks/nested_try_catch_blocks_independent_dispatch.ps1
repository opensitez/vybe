# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/nested_try_catch_blocks_independent_dispatch
$events = [System.Collections.Generic.List[string]]::new()
try {
    $events.Add("OuterTry")
    try {
        $events.Add("InnerTry")
        throw [System.FormatException]::new()
    } catch [System.FormatException] {
        $events.Add("InnerCatchFormat")
    }
    $events.Add("OuterAfterInner")
} catch {
    $events.Add("OuterCatch")
}
$res = $events -join "->"
if ($res -ne "OuterTry->InnerTry->InnerCatchFormat->OuterAfterInner") {
    Write-Host "FAIL: Nested try-catch dispatch failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
