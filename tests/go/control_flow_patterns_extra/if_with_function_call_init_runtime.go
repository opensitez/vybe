// vybe-test: go/control_flow_patterns_extra/if_with_function_call_init_runtime
// origin: languages/go/tests/go/test_control_flow_patterns_extra.rs

package main
import "fmt"
func value() int { return 7 }
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

func main() { if n := value(); n > 5 { __p(fmt.Sprint(n)) } 
__check("7")
}
