// vybe-test: go/interface_assertion_extended/typed_nil_assert_to_concrete_pointer_ok
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type widget struct { n int }
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

func main() { var p *widget
var v interface{} = p
_, ok := v.(*widget)
__p(fmt.Sprint(ok)) 
__check("true")
}
