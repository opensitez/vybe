// vybe-test: go/generics_types/generic_comparer_ordered_constraint
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
import "cmp"
type Comparer[T cmp.Ordered] interface { Less(a, b T) bool }
type IntCmp struct{}
func (IntCmp) Less(a, b int) bool { return a < b }
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

func main() { var c Comparer[int] = IntCmp{}
__p(fmt.Sprint(c.Less(2, 5))) 
__check("true")
}
