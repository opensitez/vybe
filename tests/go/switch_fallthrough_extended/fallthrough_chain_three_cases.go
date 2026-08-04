// vybe-test: go/switch_fallthrough_extended/fallthrough_chain_three_cases
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

func main() { switch 1 { case 1: __p(fmt.Sprint(1))
fallthrough
case 2: __p(fmt.Sprint(2))
fallthrough
case 3: __p(fmt.Sprint(3)) } 
__check("1\n2\n3")
}
