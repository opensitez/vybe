// vybe-test: go/generics_constraints_extended/generic_ordered_three_way_compare
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
func Compare3[T cmp.Ordered](a, b, c T) T { if cmp.Less(a, b) { return a }
if cmp.Less(b, c) { return b }
return c }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Compare3(5, 2, 8)), "2") }
