// vybe-test: go/interface_nil_comparable/generic_comparable_string_nil_map_key_lookup
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func lookup[K comparable, V any](m map[K]V, key K) bool { _, ok := m[key]
return ok }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m map[string]int
__check(fmt.Sprint(lookup(m, "missing")), "false") }
