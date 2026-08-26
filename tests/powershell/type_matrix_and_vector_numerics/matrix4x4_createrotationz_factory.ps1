# vybe-test: powershell/type_matrix_and_vector_numerics/matrix4x4_createrotationz_factory
$mat = [System.Numerics.Matrix4x4]::CreateRotationZ(0.0)
if (-not $mat.IsIdentity) { Write-Host "FAIL: Matrix4x4 CreateRotationZ(0) should be Identity"; exit 1 }
Write-Host "PASS"; exit 0
