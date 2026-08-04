# vybe-test: powershell/functions/begin_process_end
function Sum-Values {
    param([Parameter(ValueFromPipeline=$true)]$Value)
    begin { $total = 0 }
    process { $total += $Value }
    end { return $total }
}
$result = 1,2,3,4,5 | Sum-Values
if ($result -ne 15) {
    Write-Host "FAIL: expected 15, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
