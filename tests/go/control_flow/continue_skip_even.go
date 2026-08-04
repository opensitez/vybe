// vybe-test: go/control_flow/continue_skip_even
// origin: languages/go/tests/go/test_control_flow.rs

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

func main() { for i := 0; i < 6; i++ { if i % 2 == 0 { continue }
__p(fmt.Sprint(i))
} 
__check("1\n3\n5")
}
