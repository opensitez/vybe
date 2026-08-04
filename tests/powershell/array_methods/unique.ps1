# vybe-test: powershell/array_methods/unique
$arr = 1,2,2,3
if (($(($arr | Select-Object -Unique) -join ',') -eq '1,2,3')) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
EOF && printf 'array_methods: %d\n' "$(find tests/powershell/array_methods -maxdepth 1 -type f -name '*.ps1' | wc -l)"
