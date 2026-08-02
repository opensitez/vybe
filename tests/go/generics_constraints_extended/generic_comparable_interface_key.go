// vybe-test: go/generics_constraints_extended/generic_comparable_interface_key
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Count[K comparable, V any](m map[K]V) int { return len(m) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { type I interface { ~int }
__check(fmt.Sprint(Count(map[int]string{1: "a"})), "1") }
