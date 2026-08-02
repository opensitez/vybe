// vybe-test: go/generics_constraints_extended/generic_comparable_map_equal_keys
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func SameKeys[K comparable, V any](a, b map[K]V) bool { if len(a) != len(b) { return false }
for k := range a { if _, ok := b[k]; !ok { return false } }
return true }
func main() { a := map[string]int{"x": 1}
b := map[string]int{"x": 2}
fmt.Println(SameKeys(a, b)) }
