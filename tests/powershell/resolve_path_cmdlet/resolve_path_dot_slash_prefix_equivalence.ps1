# vybe-test: powershell/resolve_path_cmdlet/resolve_path_dot_slash_prefix_equivalence
# Resolving './tests' yields the exact same PathInfo as resolving 'tests'
$withDot = Resolve-Path "./tests"
$withoutDot = Resolve-Path "tests"

if ($withDot.Path -ne $withoutDot.Path) {
    Write-Host "FAIL: './tests' ($($withDot.Path)) did not match 'tests' ($($withoutDot.Path))"
    exit 1
}

Write-Host "PASS"
exit 0
