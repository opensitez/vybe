# vybe-test: powershell/string_builder_operations/indexer_character_access_and_mutation
$sb = [System.Text.StringBuilder]::new("cat")
$sb[0] = [char]'b'
if ($sb.ToString() -ne "bat" -or $sb[1] -ne [char]'a') {
    Write-Host "FAIL: Indexer character modification failed, got $($sb.ToString())"
    exit 1
}
Write-Host "PASS"
exit 0
