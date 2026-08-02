// vybe-test: go/generics_types/generic_comparer_ordered_constraint
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
import "cmp"
type Comparer[T cmp.Ordered] interface { Less(a, b T) bool }
type IntCmp struct{}
func (IntCmp) Less(a, b int) bool { return a < b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var c Comparer[int] = IntCmp{}
__check(fmt.Sprint(c.Less(2, 5)), "true") }
