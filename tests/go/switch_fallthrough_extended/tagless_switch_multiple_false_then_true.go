// vybe-test: go/switch_fallthrough_extended/tagless_switch_multiple_false_then_true
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

func main() { n := 12
switch { case n%2 == 1: __p(fmt.Sprint("odd"))
case n%3 == 0: __p(fmt.Sprint("three"))
case n%4 == 0: __p(fmt.Sprint("four"))
default: __p(fmt.Sprint("none")) } 
__check("three")
}
