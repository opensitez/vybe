# vybe-test: powershell/functions/function_output_vs_return
function Compute([int]$n) {
    "intermediate"    # this goes into the output stream
    return $n * 2
}
$all = Compute 5
# $all is an array: ["intermediate", 10]
if ($all.Count -ne 2)          { Write-Host "FAIL: count $($all.Count)"; exit 1 }
if ($all[0] -ne "intermediate") { Write-Host "FAIL: [0]"; exit 1 }
if ($all[1] -ne 10)             { Write-Host "FAIL: [1]"; exit 1 }
Write-Host "PASS"
exit 0
