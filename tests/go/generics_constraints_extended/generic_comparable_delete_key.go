// vybe-test: go/generics_constraints_extended/generic_comparable_delete_key
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Del[K comparable, V any](m map[K]V, k K) { delete(m, k) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[int]string{1: "a", 2: "b"}
Del(m, 1)
__check(fmt.Sprint(len(m)), "1") }
