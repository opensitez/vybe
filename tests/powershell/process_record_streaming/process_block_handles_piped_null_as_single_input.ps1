# vybe-test: powershell/process_record_streaming/process_block_handles_piped_null_as_single_input
# Piping $null into a function executes the process block once with $_ equal to $null
$script:nullProcessCalls = 0
$script:sawNull = $false

function NullConsumer {
    process {
        $script:nullProcessCalls++
        if ($_ -eq $null) {
            $script:sawNull = $true
        }
    }
}

$null | NullConsumer

if ($script:nullProcessCalls -ne 1) {
    Write-Host "FAIL: expected process to run 1 time for piped null, ran $($script:nullProcessCalls) times"
    exit 1
}

if (-not $script:sawNull) {
    Write-Host "FAIL: process block did not receive null in `$_"
    exit 1
}

Write-Host "PASS"
exit 0
