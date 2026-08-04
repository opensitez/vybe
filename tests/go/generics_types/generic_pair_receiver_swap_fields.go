// vybe-test: go/generics_types/generic_pair_receiver_swap_fields
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Pair[T any] struct { First, Second T }
func (p *Pair[T]) Swap() { p.First, p.Second = p.Second, p.First }
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

func main() { p := Pair[int]{First: 1, Second: 2}
p.Swap()
__p(fmt.Sprint(p.First))
__p(fmt.Sprint(p.Second)) 
__check("2\n1")
}
