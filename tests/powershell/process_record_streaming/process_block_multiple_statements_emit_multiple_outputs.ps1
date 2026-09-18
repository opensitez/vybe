# vybe-test: powershell/process_record_streaming/process_block_multiple_statements_emit_multiple_outputs
# A single process block iteration can emit multiple individual objects into the downstream pipeline
function DuplicateElements {
    process {
        "orig:$_"
        "copy:$_"
    }
}

$results = @(1, 2) | DuplicateElements

# 1 produces orig:1, copy:1; 2 produces orig:2, copy:2 => total 4 items
if ($results.Count -ne 4) {
    Write-Host "FAIL: expected 4 items from multi-output process, got $($results.Count)"
    exit 1
}

$expected = "orig:1, copy:1, orig:2, copy:2"
$actual = $results -join ", "

if ($actual -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$actual'"
    exit 1
}

Write-Host "PASS"
exit 0
