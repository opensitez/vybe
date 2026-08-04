// vybe-test: go/control_flow_patterns_extra/for_with_multiple_init_vars_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

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

func main() { for i, j := 0, 3; i < 3; i, j = i+1, j-1 { __p(fmt.Sprint(i + j)) } 
__check("3\n3\n3")
}
