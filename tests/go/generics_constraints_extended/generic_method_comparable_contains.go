// vybe-test: go/generics_constraints_extended/generic_method_comparable_contains
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Set[T comparable] struct { items []T }
func (s *Set[T]) Add(v T) { s.items = append(s.items, v) }
func (s Set[T]) Has(v T) bool { for _, x := range s.items { if x == v { return true } }
return false }
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

func main() { var st Set[int]
st.Add(7)
__p(fmt.Sprint(st.Has(7)))
__p(fmt.Sprint(st.Has(8))) 
__check("true\nfalse")
}
