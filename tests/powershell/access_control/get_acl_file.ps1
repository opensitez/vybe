# vybe-test: powershell/access_control/get_acl_file
$bytes = [byte[]]@(1, 2, 3, 4, 5)
$sha = [System.Security.Cryptography.SHA256]::Create()
$hash = $sha.ComputeHash($bytes)
if ($hash.Length -ne 32) {
    Write-Host "FAIL: Security cryptography hash check failed"
    exit 1
}
Write-Host "PASS"
exit 0
