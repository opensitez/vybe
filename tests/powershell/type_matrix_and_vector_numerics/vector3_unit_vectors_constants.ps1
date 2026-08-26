# vybe-test: powershell/type_matrix_and_vector_numerics/vector3_unit_vectors_constants
$x = [System.Numerics.Vector3]::UnitX
$y = [System.Numerics.Vector3]::UnitY
$z = [System.Numerics.Vector3]::UnitZ
if ($x.X -ne 1 -or $y.Y -ne 1 -or $z.Z -ne 1) { Write-Host "FAIL: Vector3 Unit vectors failed"; exit 1 }
Write-Host "PASS"; exit 0
