# vybe-test: powershell/process_record_streaming/process_block_without_param_block_accepts_pipeline
# A function without an explicit param() block automatically binds pipeline inputs into $_ inside process
function SimpleProcessOnly {
    process {
        "processed:$_"
    }
}

$results = @("alpha", "beta") | SimpleProcessOnly

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 items, got $($results.Count)"
    exit 1
}

if ($results[0] -ne "processed:alpha" -or $results[1] -ne "processed:beta") {
    Write-Host "FAIL: unexpected results: $($results -join ', ')"
    exit 1
}

Write-Host "PASS"
exit 0
