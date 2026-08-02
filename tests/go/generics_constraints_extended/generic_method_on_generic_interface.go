// vybe-test: go/generics_constraints_extended/generic_method_on_generic_interface
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Stringer[T any] interface { Format() string }
type Item[T any] struct { V T }
func (i Item[T]) Format() string { return fmt.Sprintf("%v", i.V) }
func Print[T any](s Stringer[T]) { __check(fmt.Sprint(s.Format()), "7") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { Print(Item[int]{V: 7}) }
