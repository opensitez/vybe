// vybe-test: go/methods/method_chaining
// origin: languages/go/tests/go/test_methods.rs

package main
import "fmt"
type Calc struct { n int }
func (c *Calc) Add(x int) *Calc { c.n += x
return c }
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

func main() { c := &Calc{n: 0}
c.Add(5).Add(3)
__p(fmt.Sprint(c.n))
__check("8")
}
