// vybe-test: go/generics_types/generic_holder_value_receiver_peek
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Holder[T any] struct { items []T }
func (h Holder[T]) Peek() T { return h.items[0] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Holder[int]{items: []int{11}}.Peek()), "11") }
