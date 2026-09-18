# vybe-test: powershell/extended_type_system/ets_add_noteproperty
# Update-TypeData adds a static NoteProperty across all instances of a type
Update-TypeData -TypeName System.DateTime -MemberType NoteProperty -MemberName FixedOrigin -Value "UTC-BASE" -Force
$dt = [DateTime]::Now

if ($dt.FixedOrigin -ne "UTC-BASE") {
    Write-Host "FAIL: expected NoteProperty value 'UTC-BASE', got: '$($dt.FixedOrigin)'"
    exit 1
}

Write-Host "PASS"
exit 0
