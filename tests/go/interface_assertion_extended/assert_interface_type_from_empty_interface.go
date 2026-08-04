// vybe-test: go/interface_assertion_extended/assert_interface_type_from_empty_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type fmtStringer interface { String() string }
type myInt int
func (m myInt) String() string { return "n" }
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

func main() { var v interface{} = myInt(3)
s, ok := v.(fmtStringer)
__p(fmt.Sprint(s.String()))
__p(fmt.Sprint(ok)) 
__check("n\ntrue")
}
