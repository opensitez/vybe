// vybe-test: go/switch_fallthrough_extended/switch_in_func_returns_via_case
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func label(n int) string { switch n { case 1: return "one"
default: return "many" } }
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

func main() { __p(fmt.Sprint(label(1)))
__p(fmt.Sprint(label(9))) 
__check("one\nmany")
}
