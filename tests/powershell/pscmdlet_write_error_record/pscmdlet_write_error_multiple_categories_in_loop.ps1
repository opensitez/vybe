# vybe-test: powershell/pscmdlet_write_error_record/pscmdlet_write_error_multiple_categories_in_loop
# Iterative calls to $PSCmdlet.WriteError with distinct ErrorCategory values preserve each category faithfully
function EmitMultiCategoryErrors {
    [CmdletBinding()]
    param()
    process {
        $categories = @(
            [System.Management.Automation.ErrorCategory]::SyntaxError,
            [System.Management.Automation.ErrorCategory]::PermissionDenied,
            [System.Management.Automation.ErrorCategory]::DeviceError
        )

        foreach ($cat in $categories) {
            $ex = [System.Exception]::new("CategoryError: $cat")
            $err = [System.Management.Automation.ErrorRecord]::new(
                $ex,
                "CatErrId_$cat",
                $cat,
                $null
            )
            $PSCmdlet.WriteError($err)
        }
    }
}

$captured = @(EmitMultiCategoryErrors 2>&1)

if ($captured.Count -ne 3) {
    Write-Host "FAIL: expected 3 category errors, got $($captured.Count)"
    exit 1
}

$expectedCategories = @("SyntaxError", "PermissionDenied", "DeviceError")
for ($i = 0; $i -lt 3; $i++) {
    $actualCat = $captured[$i].CategoryInfo.Category.ToString()
    if ($actualCat -ne $expectedCategories[$i]) {
        Write-Host "FAIL: at index ${i}, expected '$($expectedCategories[$i])', got '$actualCat'"
        exit 1
    }
}

Write-Host "PASS"
exit 0
