// vybe-test: go/interface_assertion_extended/assert_empty_interface_from_named_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type speaker interface { talk() string }
type bot struct{}
func (bot) talk() string { return "hi" }
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

func main() { var s speaker = bot{}
var v interface{} = s
b, ok := v.(bot)
__p(fmt.Sprint(b.talk()))
__p(fmt.Sprint(ok)) 
__check("hi\ntrue")
}
