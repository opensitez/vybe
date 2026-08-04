// vybe-test: go/interfaces/struct_method_modify_returns_new
// origin: languages/go/tests/go/test_interfaces.rs

package main
import "fmt"
type Counter struct { N int } func (c Counter) Inc() Counter { return Counter{N: c.N + 1} } var __buf string

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

func main() { c := Counter{N: 0}
c = c.Inc()
c = c.Inc()
__p(fmt.Sprint(c.N))
__check("2")
}
