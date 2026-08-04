// vybe-test: go/generics_types/generic_container_interface_impl
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Container[T any] interface { Add(T)
Size() int }
type SliceBox[T any] struct { items []T }
func (s *SliceBox[T]) Add(v T) { s.items = append(s.items, v) }
func (s *SliceBox[T]) Size() int { return len(s.items) }
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

func main() { var c Container[int] = &SliceBox[int]{}
c.Add(5)
c.Add(8)
__p(fmt.Sprint(c.Size())) 
__check("2")
}
