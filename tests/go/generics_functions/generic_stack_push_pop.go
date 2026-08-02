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
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var st Stack[int]
st.Push(5)
__check(fmt.Sprint(st.Pop()), "5") }
