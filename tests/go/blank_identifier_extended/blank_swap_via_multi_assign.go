// vybe-test: go/blank_identifier_extended/blank_swap_via_multi_assign
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
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

func main() { a, b := 1, 2
a, b = b, a
_, _ = a, b
__p(fmt.Sprint(a))
__p(fmt.Sprint(b)) 
__check("2\n1")
}
