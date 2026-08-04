// vybe-test: go/generics_types/generic_set_comparable_membership
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Set[T comparable] struct { m map[T]struct{} }
func (s *Set[T]) Add(v T) { if s.m == nil { s.m = make(map[T]struct{}) }
s.m[v] = struct{}{} }
func (s *Set[T]) Has(v T) bool { _, ok := s.m[v]
return ok }
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

func main() { var s Set[int]
s.Add(3)
__p(fmt.Sprint(s.Has(3)))
__p(fmt.Sprint(s.Has(4))) 
__check("true\nfalse")
}
