// vybe-test: go/generics_constraints_extended/generic_method_value_receiver_get
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Cell[T any] struct { V T }
func (c Cell[T]) Get() T { return c.V }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Cell[int]{V: 42}.Get()), "42") }
