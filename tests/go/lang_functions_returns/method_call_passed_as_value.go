// vybe-test: go/lang_functions_returns/method_call_passed_as_value
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
type S struct{}
func (S) ID() int { return 7 }
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

func main() { var s S
__p(fmt.Sprint(s.ID())) 
__check("7")
}
