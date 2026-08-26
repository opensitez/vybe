# vybe-test: powershell/classes_static_constructors/static_constructor_populates_bitarray_flags
class BitFlags {
    static [System.Collections.BitArray]$DefaultMask
    static BitFlags() {
        [BitFlags]::DefaultMask = [System.Collections.BitArray]::new(@($true, $false, $true, $false))
    }
}
if ([BitFlags]::DefaultMask[0] -ne $true -or [BitFlags]::DefaultMask[1] -ne $false) {
    Write-Host "FAIL: Static BitArray flags failed"
    exit 1
}
Write-Host "PASS"
exit 0
