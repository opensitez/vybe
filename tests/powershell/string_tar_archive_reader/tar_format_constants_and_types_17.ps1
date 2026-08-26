# vybe-test: powershell/string_tar_archive_reader/tar_format_constants_and_types_17
$entryType = [System.Formats.Tar.TarEntryType]::RegularFile
if ([int]$entryType -ne 48) { Write-Host "FAIL: TarEntryType RegularFile expected 48"; exit 1 }
Write-Host "PASS"; exit 0
