// vybe-test: go/generics_constraints_extended/generic_comparable_map_keys_to_slice
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func KeyList[K comparable, V any](m map[K]V) []K { keys := make([]K, 0, len(m))
for k := range m { keys = append(keys, k) }
return keys }
func main() { fmt.Println(len(KeyList(map[int]bool{1: true, 2: false}))) }
