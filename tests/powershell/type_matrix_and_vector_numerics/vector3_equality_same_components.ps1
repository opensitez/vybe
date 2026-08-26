# vybe-test: powershell/type_matrix_and_vector_numerics/vector3_equality_same_components
$v1 = [System.Numerics.Vector3]::new(1.0, 2.0, 3.0)
$v2 = [System.Numerics.Vector3]::new(1.0, 2.0, 3.0)
if (-not $v1.Equals($v2)) { Write-Host "FAIL: Vector3 Equals failed"; exit 1 }
Write-Host "PASS"; exit 0
