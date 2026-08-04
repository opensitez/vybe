// vybe-test: go/interface_assertion_extended/comma_ok_assert_named_interface_from_empty_false
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type speaker interface { talk() string }
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

func main() { var v interface{} = 1
_, ok := v.(speaker)
__p(fmt.Sprint(ok)) 
__check("false")
}
