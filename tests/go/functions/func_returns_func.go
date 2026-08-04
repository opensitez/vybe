// vybe-test: go/functions/func_returns_func
// origin: languages/go/tests/go/test_functions.rs

package main
import "fmt"
func adder(n int) func(int) int { return func(x int) int { return x + n } } var __buf string

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

func main() { add10 := adder(10)
__p(fmt.Sprint(add10(5)))
__check("15")
}
