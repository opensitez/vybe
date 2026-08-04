// vybe-test: go/higher_order_functions/compose_two_functions
// origin: languages/go/tests/go/test_higher_order_functions.rs

package main
import "fmt"
func compose(f func(int) int, g func(int) int) func(int) int { return func(x int) int { return f(g(x)) } }
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

func main() { double := func(x int) int { return x * 2 }
inc := func(x int) int { return x + 1 }
h := compose(double, inc)
__p(fmt.Sprint(h(3))) 
__check("8")
}
