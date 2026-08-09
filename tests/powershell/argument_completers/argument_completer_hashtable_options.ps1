# vybe-test: powershell/argument_completers/argument_completer_hashtable_options
$optTable = @{ Environment = @("Dev", "Staging", "Prod") }
$completer = {
    param($key)
    $script:optTable[$key]
}
$res = @(&$completer "Environment")
if ($res.Count -ne 3 -or $res[2] -ne "Prod") {
    Write-Host "FAIL: hashtable argument completer expected Dev, Staging, Prod"
    exit 1
}
Write-Host "PASS"
exit 0
