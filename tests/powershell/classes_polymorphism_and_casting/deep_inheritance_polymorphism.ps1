# vybe-test: powershell/classes_polymorphism_and_casting/deep_inheritance_polymorphism
class LevelA { [string]WhoAmI() { return "A" } }
class LevelB : LevelA { [string]WhoAmI() { return "B" } }
class LevelC : LevelB { [string]WhoAmI() { return "C" } }
class LevelD : LevelC { [string]WhoAmI() { return "D" } }
[LevelA]$top = [LevelD]::new()
if ($top.WhoAmI() -ne "D") {
    Write-Host "FAIL: Deep inheritance dispatch expected 'D', got '$($top.WhoAmI())'"
    exit 1
}
Write-Host "PASS"
exit 0
