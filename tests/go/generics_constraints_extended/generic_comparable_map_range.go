// vybe-test: go/generics_constraints_extended/generic_comparable_map_range
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
func SumValues[K comparable, V ~int](m map[K]V) int { s := 0
for _, v := range m { s += int(v) }
return s }
func main() { _ = SumValues(map[string]int8{"a": 1}) }
