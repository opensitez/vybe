// vybe-test: go/switch_fallthrough_extended/switch_on_string_multi_case
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

func main() { switch "a" { case "a", "b": __p(fmt.Sprint("ab"))
default: __p(fmt.Sprint("x")) } 
__check("ab")
}
