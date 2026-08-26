# vybe-test: powershell/type_matrix_and_vector_numerics/quaternion_identity_static_property
$q = [System.Numerics.Quaternion]::Identity
if (-not $q.IsIdentity -or $q.W -ne 1.0 -or $q.X -ne 0.0) { Write-Host "FAIL: Quaternion Identity failed"; exit 1 }
Write-Host "PASS"; exit 0
