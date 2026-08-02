// vybe-test: go/generics_constraints_extended/generic_method_pointer_receiver_set
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Cell[T any] struct { V T }
func (c *Cell[T]) Set(v T) { c.V = v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := Cell[string]{}
c.Set("ok")
__check(fmt.Sprint(c.V), "ok") }
