# vybe-test: powershell/type_matrix_and_vector_numerics/matrix3x2_createtranslation_factory
$mat = [System.Numerics.Matrix3x2]::CreateTranslation(10.0, 20.0)
if ($mat.M31 -ne 10.0 -or $mat.M32 -ne 20.0) { Write-Host "FAIL: Matrix3x2 CreateTranslation failed"; exit 1 }
Write-Host "PASS"; exit 0
