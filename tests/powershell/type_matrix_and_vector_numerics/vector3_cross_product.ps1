# vybe-test: powershell/type_matrix_and_vector_numerics/vector3_cross_product
$v1 = [System.Numerics.Vector3]::UnitX
$v2 = [System.Numerics.Vector3]::UnitY
$cross = [System.Numerics.Vector3]::Cross($v1, $v2)
if ($cross.X -ne 0 -or $cross.Y -ne 0 -or $cross.Z -ne 1) { Write-Host "FAIL: Vector3 Cross expected UnitZ, got $cross"; exit 1 }
Write-Host "PASS"; exit 0
