# vybe-test: powershell/type_matrix_and_vector_numerics/matrix4x4_determinant_calculation
$id = [System.Numerics.Matrix4x4]::Identity
$det = $id.GetDeterminant()
if ($det -ne 1.0) { Write-Host "FAIL: Matrix4x4 Determinant expected 1, got $det"; exit 1 }
Write-Host "PASS"; exit 0
