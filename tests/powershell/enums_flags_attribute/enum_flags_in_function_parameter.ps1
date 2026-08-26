# vybe-test: powershell/enums_flags_attribute/enum_flags_in_function_parameter
[System.FlagsAttribute()]
enum RunMode {
    Debug = 1
    Verbose = 2
    Silent = 4
}
function Invoke-Runner([RunMode]$mode) {
    if ($mode.HasFlag([RunMode]::Debug)) { return "DebugMode" }
    return "NormalMode"
}
$r1 = Invoke-Runner ([RunMode]::Debug -bor [RunMode]::Verbose)
$r2 = Invoke-Runner ([RunMode]::Silent)
if ($r1 -ne "DebugMode" -or $r2 -ne "NormalMode") {
    Write-Host "FAIL: Flags enum in function parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
