// vybe-test: go/method_sets_pointer_value/value_method_on_nonaddressable_temp_via_interface_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type speaker interface { say() string }
type bot struct{}
func (b bot) say() string { return "beep" }
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
__p(fmt.Sprint(s.say())) 
__check("beep")
}
