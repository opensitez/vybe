// vybe-test: go/functions/function_multiple_returns
// origin: languages/go/tests/go/test_functions.rs

package main
import "fmt"
func divmod(a int, b int) (int, int) { return a / b, a % b } var __buf string

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

func main() { q, r := divmod(17, 5)
__p(fmt.Sprint(q))
__p(fmt.Sprint(r))
__check("3\n2")
}
