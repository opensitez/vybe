// vybe-test: go/interface_nil_comparable/comparable_generic_map_key_compile
// origin: languages/go/tests/go/test_interface_nil_comparable.rs
// vybe-test-mode: compile

package main
func keys[K comparable, V any](m map[K]V) []K { result := make([]K, 0)
for k := range m { result = append(result, k) }
return result }
func main() { _ = keys(map[int]string{}) }
