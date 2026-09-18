# vybe-test: powershell/loops/for_loop_omitted_sections
# A for loop supports omitting the initializer section when the variable is pre-initialized
$accumulator = 0
$counter = 5

for (; $counter -lt 10; $counter++) {
    $accumulator += $counter
}

# Sum of 5 + 6 + 7 + 8 + 9 = 35; final counter = 10
if ($accumulator -ne 35) {
    Write-Host "FAIL: expected sum 35, got $accumulator"
    exit 1
}

if ($counter -ne 10) {
    Write-Host "FAIL: expected counter 10, got $counter"
    exit 1
}

Write-Host "PASS"
exit 0
