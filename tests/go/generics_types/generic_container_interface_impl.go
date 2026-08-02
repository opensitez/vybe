// vybe-test: go/generics_types/generic_container_interface_impl
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Container[T any] interface { Add(T)
Size() int }
type SliceBox[T any] struct { items []T }
func (s *SliceBox[T]) Add(v T) { s.items = append(s.items, v) }
func (s *SliceBox[T]) Size() int { return len(s.items) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var c Container[int] = &SliceBox[int]{}
c.Add(5)
c.Add(8)
__check(fmt.Sprint(c.Size()), "2") }
