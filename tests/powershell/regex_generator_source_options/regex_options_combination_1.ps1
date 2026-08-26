# vybe-test: powershell/regex_generator_source_options/regex_options_combination_1
$opt = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled
$m = [System.Text.RegularExpressions.Regex]::Match("ITEM_1", "item_1", $opt)
if (-not $m.Success) { Write-Host "FAIL: Regex match with options failed"; exit 1 }
Write-Host "PASS"; exit 0
