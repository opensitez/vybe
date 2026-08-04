// vybe-test: go/generics_functions/generic_stack_push_pop
// origin: languages/go/tests/go/test_generics_functions.rs

package main
import "fmt"
type Stack[T any] struct { items []T }
func (s *Stack[T]) Push(v T) { s.items = append(s.items, v) }
func (s *Stack[T]) Pop() T { n := len(s.items)-1
v := s.items[n]
s.items = s.items[:n]
return v }
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

func main() { var st Stack[int]
st.Push(5)
__p(fmt.Sprint(st.Pop())) 
__check("5")
}
