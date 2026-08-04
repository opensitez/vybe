// vybe-test: go/generics_constraints_extended/generic_method_chained_on_type_param
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Num[T ~int] struct { V T }
func (n Num[T]) Inc() Num[T] { n.V++
return n }
func (n Num[T]) Value() T { return n.V }
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

func main() { __p(fmt.Sprint(Num[int]{V: 1}.Inc().Value())) 
__check("2")
}
