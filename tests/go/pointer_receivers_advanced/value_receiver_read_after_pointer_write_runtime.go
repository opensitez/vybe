// vybe-test: go/pointer_receivers_advanced/value_receiver_read_after_pointer_write_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type note struct { text string }
func (n *note) set(v string) { n.text = v }
func (n note) read() string { return n.text }
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

func main() { value := note{text: "a"}
value.set("b")
__p(fmt.Sprint(value.read()))
__check("b")
}
