# vybe-test: powershell/type_matrix_and_vector_numerics/matrix3x2_identity_static_property
$id = [System.Numerics.Matrix3x2]::Identity
if (-not $id.IsIdentity -or $id.M11 -ne 1.0 -or $id.M22 -ne 1.0) { Write-Host "FAIL: Matrix3x2 Identity failed"; exit 1 }
Write-Host "PASS"; exit 0
