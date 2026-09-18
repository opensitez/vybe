# vybe-test: powershell/culture_and_globalization_cmdlets/culture_arabic_textinfo_right_to_left
# The Arabic (ar-SA) culture TextInfo indicates right-to-left text direction
$ar = Get-Culture -Name "ar-SA"

if (-not $ar.TextInfo.IsRightToLeft) {
    Write-Host "FAIL: expected IsRightToLeft=true for Arabic culture, got: $($ar.TextInfo.IsRightToLeft)"
    exit 1
}

Write-Host "PASS"
exit 0
