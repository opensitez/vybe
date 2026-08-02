// vybe-test: go/generics_constraints_extended/generic_ordered_min_string
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
func Min[T cmp.Ordered](a, b T) T { if cmp.Less(a, b) { return a }
return b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Min("zebra", "apple")), "apple") }
