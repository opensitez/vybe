# vybe-test: powershell/invoke_expression_cmdlet/iex_pscustomobject_construction
# Invoke-Expression can construct a PSCustomObject from a literal string representation
$obj = Invoke-Expression '[pscustomobject]@{Id=1; Title="Test"}'

if ($obj.Id -ne 1 -or $obj.Title -ne "Test") {
    Write-Host "FAIL: PSCustomObject mismatch, Id=$($obj.Id), Title='$($obj.Title)'"
    exit 1
}

Write-Host "PASS"
exit 0
