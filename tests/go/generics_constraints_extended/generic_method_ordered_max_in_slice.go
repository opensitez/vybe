// vybe-test: go/generics_constraints_extended/generic_method_ordered_max_in_slice
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
type Stats[T cmp.Ordered] struct { data []T }
func (s Stats[T]) Max() T { m := s.data[0]
for _, v := range s.data[1:] { if cmp.Less(m, v) { m = v } }
return m }
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

func main() { __p(fmt.Sprint(Stats[int]{data: []int{1, 9, 3}}.Max())) 
__check("9")
}
