// vybe-test: go/generics_types/generic_list_append_method
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type List[T any] struct { items []T }
func (l *List[T]) Append(v T) { l.items = append(l.items, v) }
func (l List[T]) At(i int) T { return l.items[i] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var l List[string]
l.Append("a")
l.Append("b")
__check(fmt.Sprint(l.At(1)), "b") }
