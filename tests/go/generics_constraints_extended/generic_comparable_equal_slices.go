// vybe-test: go/generics_constraints_extended/generic_comparable_equal_slices
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Eq[T comparable](a, b T) bool { return a == b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Eq(1, 1)), "true")
__check(fmt.Sprint(Eq(1, 2)), "false") }
