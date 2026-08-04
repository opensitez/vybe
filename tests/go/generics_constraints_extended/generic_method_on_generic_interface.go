// vybe-test: go/generics_constraints_extended/generic_method_on_generic_interface
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Stringer[T any] interface { Format() string }
type Item[T any] struct { V T }
func (i Item[T]) Format() string { return fmt.Sprintf("%v", i.V) }
func Print[T any](s Stringer[T]) { __p(fmt.Sprint(s.Format())) }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { Print(Item[int]{V: 7}) 
__check("7")
}
