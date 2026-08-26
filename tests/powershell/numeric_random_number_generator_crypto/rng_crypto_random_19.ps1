# vybe-test: powershell/numeric_random_number_generator_crypto/rng_crypto_random_19
$rndInt = [System.Security.Cryptography.RandomNumberGenerator]::GetInt32(1, 100)
if ($rndInt -lt 1 -or $rndInt -ge 100) { Write-Host "FAIL: GetInt32 bounds failed"; exit 1 }
Write-Host "PASS"; exit 0
