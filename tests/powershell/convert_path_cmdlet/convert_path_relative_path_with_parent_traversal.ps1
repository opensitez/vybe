# vybe-test: powershell/convert_path_cmdlet/convert_path_relative_path_with_parent_traversal
# Relative paths containing internal '..' traversals are normalized to their canonical path
$traversed = Convert-Path "tests/../tests"
$canonical = Convert-Path "tests"

if ($traversed -ne $canonical) {
    Write-Host "FAIL: internal parent traversal did not normalize to canonical path, expected '$canonical', got: '$traversed'"
    exit 1
}

Write-Host "PASS"
exit 0
