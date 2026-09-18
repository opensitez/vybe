# vybe-test: powershell/resolve_path_cmdlet/resolve_path_relative_parent_traversal
# Paths containing internal '..' traversals are normalized to their canonical PathInfo representation
$traversed = Resolve-Path "tests/../tests"
$canonical = Resolve-Path "tests"

if ($traversed.Path -ne $canonical.Path) {
    Write-Host "FAIL: internal parent traversal did not resolve to canonical path, expected '$($canonical.Path)', got: '$($traversed.Path)'"
    exit 1
}

Write-Host "PASS"
exit 0
