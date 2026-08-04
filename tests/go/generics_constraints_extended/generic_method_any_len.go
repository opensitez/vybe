// vybe-test: go/generics_constraints_extended/generic_method_any_len
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Bag[T any] struct { items []T }
func (b Bag[T]) Len() int { return len(b.items) }
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

func main() { __p(fmt.Sprint(Bag[string]{items: []string{"a", "b"}}.Len())) 
__check("2")
}
