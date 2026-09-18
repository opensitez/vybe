# vybe-test: powershell/process_record_streaming/process_block_does_not_execute_on_empty_pipeline
# Piping an empty array into a function executes begin and end, but skips the process block completely
$script:beginRan = $false
$script:processRan = $false
$script:endRan = $false

function LifecycleAudit {
    begin {
        $script:beginRan = $true
    }
    process {
        $script:processRan = $true
    }
    end {
        $script:endRan = $true
    }
}

@() | LifecycleAudit

if (-not $script:beginRan) {
    Write-Host "FAIL: begin block did not run on empty array"
    exit 1
}

if ($script:processRan) {
    Write-Host "FAIL: process block should not run when pipeline input is empty array"
    exit 1
}

if (-not $script:endRan) {
    Write-Host "FAIL: end block did not run on empty array"
    exit 1
}

Write-Host "PASS"
exit 0
