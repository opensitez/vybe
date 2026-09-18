# vybe-test: powershell/convert_path_cmdlet/convert_path_output_type_is_string
# Unlike Resolve-Path which emits PathInfo objects, Convert-Path returns raw [string]
$res = Convert-Path "."

if (-not ($res -is [string])) {
    Write-Host "FAIL: expected return type [string], got: $($res.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
