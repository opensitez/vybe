// vybe-test: go/methods/method_pointer_receiver_modify
// origin: languages/go/tests/go/test_methods.rs

package main
import "fmt"
type Counter struct { val int }
func (c *Counter) Add(n int) { c.val += n }
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

func main() { c := Counter{val: 0}
c.Add(5)
__p(fmt.Sprint(c.val))
__check("5")
}
