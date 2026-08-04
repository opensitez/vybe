// vybe-test: go/blank_identifier_extended/blank_discard_func_call_returns
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func divmod(a int, b int) (int, int) { return a / b, a % b }
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

func main() { q, _ := divmod(10, 3)
__p(fmt.Sprint(q)) 
__check("3")
}
