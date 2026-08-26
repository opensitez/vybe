# vybe-test: powershell/string_stringbuilder_chunk_enumerator/stringbuilder_chunk_operation_20
$sb = [System.Text.StringBuilder]::new("ChunkData_20")
$str = $sb.ToString()
if ($str -ne "ChunkData_20") { Write-Host "FAIL: StringBuilder failed"; exit 1 }
Write-Host "PASS"; exit 0
