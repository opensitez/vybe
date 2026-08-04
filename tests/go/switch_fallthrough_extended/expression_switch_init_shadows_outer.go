// vybe-test: go/switch_fallthrough_extended/expression_switch_init_shadows_outer
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

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

func main() { n := 1
switch n := 2; n { case 2: __p(fmt.Sprint(n))
default: __p(fmt.Sprint(0)) }
__p(fmt.Sprint(n)) 
__check("2\n1")
}
