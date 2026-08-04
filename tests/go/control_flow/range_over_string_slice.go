// vybe-test: go/control_flow/range_over_string_slice
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

func main() { langs := []string{"go","rust","c"}
for _, l := range langs { __p(fmt.Sprint(l))
} 
__check("go\nrust\nc")
}
