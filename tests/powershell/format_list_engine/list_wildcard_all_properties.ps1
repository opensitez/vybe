# vybe-test: powershell/format_list_engine/list_wildcard_all_properties
$entry = [pscustomobject]@{
    PropAlpha = "AlphaVal"
    PropBeta  = "BetaVal"
    PropGamma = "GammaVal"
}

# In Format-List, -Property * selects all note properties on the object
$output = $entry | Format-List -Property * | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "PropAlpha\s*:\s*AlphaVal" -and 
          $output -match "PropBeta\s*:\s*BetaVal" -and 
          $output -match "PropGamma\s*:\s*GammaVal")) {
    Write-Host "FAIL: wildcard * did not render all properties: $output"
    exit 1
}

Write-Host "PASS"
exit 0
