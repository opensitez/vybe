# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_block_with_pipeline_input
function Safe-ParsePipe {
    process {
        try {
            [int]::Parse($_)
        } catch [System.FormatException] {
            0
        }
    }
}
$res = @("10", "bad", "30" | Safe-ParsePipe)
if ($res.Length -ne 3 -or $res[0] -ne 10 -or $res[1] -ne 0 -or $res[2] -ne 30) {
    Write-Host "FAIL: Typed catch block with pipeline stream failed"
    exit 1
}
Write-Host "PASS"
exit 0
