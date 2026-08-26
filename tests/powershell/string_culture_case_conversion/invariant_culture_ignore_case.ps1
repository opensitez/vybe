# vybe-test: powershell/string_culture_case_conversion/invariant_culture_ignore_case
$cmp = [System.StringComparer]::InvariantCultureIgnoreCase
$res = $cmp.Compare("file.txt", "FILE.TXT")
if ($res -ne 0) {
    Write-Host "FAIL: InvariantCultureIgnoreCase compare expected 0, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
