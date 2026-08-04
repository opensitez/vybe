# vybe-test: powershell/enums/enum_in_switch
enum Season { Spring; Summer; Autumn; Winter }
function Describe-Season([Season]$s) {
    switch ($s) {
        ([Season]::Spring) { return "warm and rainy" }
        ([Season]::Summer) { return "hot and sunny" }
        ([Season]::Autumn) { return "cool and windy" }
        ([Season]::Winter) { return "cold and snowy" }
    }
}
if ((Describe-Season ([Season]::Summer)) -ne "hot and sunny") { Write-Host "FAIL: Summer"; exit 1 }
if ((Describe-Season ([Season]::Winter)) -ne "cold and snowy") { Write-Host "FAIL: Winter"; exit 1 }
Write-Host "PASS"
exit 0
