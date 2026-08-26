# vybe-test: powershell/type_matrix_and_vector_numerics/matrix4x4_identity_static_property
$id = [System.Numerics.Matrix4x4]::Identity
if (-not $id.IsIdentity -or $id.M11 -ne 1.0 -or $id.M44 -ne 1.0) { Write-Host "FAIL: Matrix4x4 Identity failed"; exit 1 }
Write-Host "PASS"; exit 0
