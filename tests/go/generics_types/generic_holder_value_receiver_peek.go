// vybe-test: go/generics_types/generic_holder_value_receiver_peek
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Holder[T any] struct { items []T }
func (h Holder[T]) Peek() T { return h.items[0] }
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

func main() { __p(fmt.Sprint(Holder[int]{items: []int{11}}.Peek())) 
__check("11")
}
