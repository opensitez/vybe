// vybe-test: go/generics_functions/generic_comparable_map_key
// origin: languages/go/tests/go/test_generics_functions.rs
// vybe-test-mode: compile

package main
func Keys[K comparable, V any](m map[K]V) []K { keys := make([]K, 0, len(m))
for k := range m { keys = append(keys, k) }
return keys }
func main() { _ = Keys(map[int]string{1:"a"}) }
