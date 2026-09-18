# vybe-test: powershell/process_record_streaming/process_block_multiple_pipeline_inputs_chained
# Multiple custom streaming functions chained together process and transform items sequentially
function AddFive {
    process {
        $_ + 5
    }
}

function MultiplyByTwo {
    process {
        $_ * 2
    }
}

function FormatStringResult {
    process {
        "val:$_"
    }
}

$results = @(1..3 | AddFive | MultiplyByTwo | FormatStringResult)

# (1+5)*2 = 12; (2+5)*2 = 14; (3+5)*2 = 16
$expected = "val:12, val:14, val:16"
$actual = $results -join ", "

if ($actual -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$actual'"
    exit 1
}

Write-Host "PASS"
exit 0
