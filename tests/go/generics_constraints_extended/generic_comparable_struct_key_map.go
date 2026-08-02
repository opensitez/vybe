// vybe-test: go/generics_constraints_extended/generic_comparable_struct_key_map
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Key struct { ID int }
func Has[K comparable, V any](m map[K]V, k K) bool { _, ok := m[k]
return ok }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { k := Key{ID: 1}
__check(fmt.Sprint(Has(map[Key]string{k: "v"}, k)), "true") }
