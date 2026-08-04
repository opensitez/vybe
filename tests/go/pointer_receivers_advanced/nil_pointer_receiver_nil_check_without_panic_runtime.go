// vybe-test: go/pointer_receivers_advanced/nil_pointer_receiver_nil_check_without_panic_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type node struct { id int }
func (n *node) absent() bool { return n == nil }
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

func main() { var value *node
__p(fmt.Sprint(value.absent()))
__check("true")
}
