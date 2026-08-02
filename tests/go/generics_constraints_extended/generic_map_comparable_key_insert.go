// vybe-test: go/generics_constraints_extended/generic_map_comparable_key_insert
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Put[K comparable, V any](m map[K]V, k K, v V) { m[k] = v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[rune]int{}
Put(m, 'x', 1)
__check(fmt.Sprint(m['x']), "1") }
