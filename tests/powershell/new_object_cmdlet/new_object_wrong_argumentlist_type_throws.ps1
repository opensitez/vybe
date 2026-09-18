# vybe-test: powershell/new_object_cmdlet/new_object_wrong_argumentlist_type_throws
# Passing an incompatible -ArgumentList value for a typed constructor throws a MethodException
$threw = $false
try {
    # DateTime constructor expects integer year; passing a string that cannot be parsed throws
    New-Object System.DateTime -ArgumentList "not-a-year" -ErrorAction Stop
} catch {
    $threw = $true
}

if (-not $threw) {
    Write-Host "FAIL: incompatible ArgumentList did not throw an exception"
    exit 1
}

Write-Host "PASS"
exit 0
