# vybe-test: powershell/type_matrix_and_vector_numerics/matrix3x2_createscale_factory
$mat = [System.Numerics.Matrix3x2]::CreateScale(2.0, 3.0)
if ($mat.M11 -ne 2.0 -or $mat.M22 -ne 3.0) { Write-Host "FAIL: Matrix3x2 CreateScale failed"; exit 1 }
Write-Host "PASS"; exit 0
