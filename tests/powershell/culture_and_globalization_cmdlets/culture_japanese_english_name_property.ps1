# vybe-test: powershell/culture_and_globalization_cmdlets/culture_japanese_english_name_property
# The EnglishName property of the Japanese (ja-JP) culture contains the word 'Japanese'
$jp = Get-Culture -Name "ja-JP"

if ($jp.EnglishName -notmatch "Japanese") {
    Write-Host "FAIL: expected 'Japanese' in EnglishName, got: '$($jp.EnglishName)'"
    exit 1
}

Write-Host "PASS"
exit 0
