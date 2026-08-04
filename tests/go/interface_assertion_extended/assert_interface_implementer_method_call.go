// vybe-test: go/interface_assertion_extended/assert_interface_implementer_method_call
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type greeter interface { greet() string }
type hi struct{}
func (hi) greet() string { return "yo" }
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

func main() { var v interface{} = hi{}
if g, ok := v.(greeter); ok { __p(fmt.Sprint(g.greet())) } else { __p(fmt.Sprint("no")) } 
__check("yo")
}
