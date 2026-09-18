# vybe-test: powershell/process_record_streaming/process_block_receives_current_item_in_psitem
# Inside a process block, $PSItem functions as an identical alias to $_
$script:doubled = @()

function DoubleViaPSItem {
    process {
        $script:doubled += ($PSItem * 2)
    }
}

@(1, 2, 3, 4) | DoubleViaPSItem

$expected = @(2, 4, 6, 8)
for ($i = 0; $i -lt 4; $i++) {
    if ($script:doubled[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected $($expected[$i]), got $($script:doubled[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
