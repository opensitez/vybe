# vybe-test: powershell/regex_match_evaluator_delegate/regex_evaluator_case_8
$eval = [System.Text.RegularExpressions.MatchEvaluator]{ param($m) return "[${m}]" }
$res = [regex]::Replace("word_8", "word_8", $eval)
if ($res -ne "[word_8]") { Write-Host "FAIL: MatchEvaluator failed"; exit 1 }
Write-Host "PASS"; exit 0
