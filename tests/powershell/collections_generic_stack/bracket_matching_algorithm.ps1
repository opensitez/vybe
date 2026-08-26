# vybe-test: powershell/collections_generic_stack/bracket_matching_algorithm
$s = [System.Collections.Generic.Stack[char]]::new()
$expr = "({[]})"
$matched = $true
foreach ($ch in $expr.ToCharArray()) {
    if ($ch -eq [char]'(' -or $ch -eq [char]'{' -or $ch -eq [char]'[') {
        $s.Push($ch)
    } elseif ($ch -eq [char]')') {
        if ($s.Count -eq 0 -or $s.Pop() -ne [char]'(') { $matched = $false; break }
    } elseif ($ch -eq [char]'}') {
        if ($s.Count -eq 0 -or $s.Pop() -ne [char]'{') { $matched = $false; break }
    } elseif ($ch -eq [char]']') {
        if ($s.Count -eq 0 -or $s.Pop() -ne [char]'[') { $matched = $false; break }
    }
}
if (-not $matched -or $s.Count -ne 0) {
    Write-Host "FAIL: Bracket matching using stack failed"
    exit 1
}
Write-Host "PASS"
exit 0
