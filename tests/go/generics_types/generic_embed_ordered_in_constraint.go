// vybe-test: go/generics_types/generic_embed_ordered_in_constraint
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
import "cmp"
type Ordered = cmp.Ordered
type Sorter[T Ordered] struct{}
func (Sorter[T]) IsLess(a, b T) bool { return a < b }
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

func main() { __p(fmt.Sprint(Sorter[int]{}.IsLess(1, 3))) 
__check("true")
}
