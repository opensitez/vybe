// vybe-test: go/defer_panic_recover_extra/defer_method_on_pointer_receiver_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
type counter struct { n int }
func (c *counter) show() { __p(fmt.Sprint(c.n)) }
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

func main() { value := &counter{n: 6}
defer value.show()
value.n = 9
__check("9")
}
