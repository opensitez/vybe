# vybe-test: powershell/interpolation/array_interpolation_uses_ofs_separator
# When an array is interpolated in a double-quoted string, elements are joined using the $OFS variable
$elements = @("red", "green", "blue")

# Default $OFS behaves as single space
$defaultJoined = "$elements"
if ($defaultJoined -ne "red green blue") {
    Write-Host "FAIL: default array interpolation expected 'red green blue', got '$defaultJoined'"
    exit 1
}

# Changing $OFS changes array interpolation delimiter
$previousOfs = $OFS
try {
    $OFS = "::"
    $customJoined = "$elements"

    if ($customJoined -ne "red::green::blue") {
        Write-Host "FAIL: custom OFS interpolation expected 'red::green::blue', got '$customJoined'"
        exit 1
    }
} finally {
    $OFS = $previousOfs
}

Write-Host "PASS"
exit 0
