// vybe-test: go/generics_constraints_extended/generic_ordered_max_float64
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
func Max[T cmp.Ordered](a, b T) T { if cmp.Less(a, b) { return b }
return a }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Max(1.5, 2.5)), "2.5") }
