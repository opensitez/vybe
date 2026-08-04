# vybe-test: powershell/enums/enum_explicit_values
enum HttpStatus {
    OK       = 200
    NotFound = 404
    Error    = 500
}
if ([int][HttpStatus]::OK       -ne 200) { Write-Host "FAIL: OK";       exit 1 }
if ([int][HttpStatus]::NotFound -ne 404) { Write-Host "FAIL: NotFound"; exit 1 }
if ([int][HttpStatus]::Error    -ne 500) { Write-Host "FAIL: Error";    exit 1 }
$status = [HttpStatus]200
if ($status -ne [HttpStatus]::OK) { Write-Host "FAIL: cast from int"; exit 1 }
Write-Host "PASS"
exit 0
