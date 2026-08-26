# vybe-test: powershell/enums_flags_attribute/enum_flags_all_flags_composite_constant
[System.FlagsAttribute()]
enum HttpMethods {
    None    = 0
    Get     = 1
    Post    = 2
    Put     = 4
    Delete  = 8
    All     = 15 # 1 + 2 + 4 + 8
}
$all = [HttpMethods]::All
if (-not $all.HasFlag([HttpMethods]::Get) -or -not $all.HasFlag([HttpMethods]::Delete)) {
    Write-Host "FAIL: Composite 'All' constant failed"
    exit 1
}
Write-Host "PASS"
exit 0
