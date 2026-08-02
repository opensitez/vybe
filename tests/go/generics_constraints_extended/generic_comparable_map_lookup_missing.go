// vybe-test: go/generics_constraints_extended/generic_comparable_map_lookup_missing
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Get[K comparable, V any](m map[K]V, k K) (V, bool) { v, ok := m[k]
return v, ok }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { _, ok := Get(map[int]string{1: "a"}, 9)
__check(fmt.Sprint(ok), "false") }
