// vybe-test: go/generics_constraints_extended/generic_ordered_three_way_compare
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
func Compare3[T cmp.Ordered](a, b, c T) T { if cmp.Less(a, b) { return a }
if cmp.Less(b, c) { return b }
return c }
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

func main() { __p(fmt.Sprint(Compare3(5, 2, 8))) 
__check("2")
}
