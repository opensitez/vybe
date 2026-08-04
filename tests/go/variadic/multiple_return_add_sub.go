// vybe-test: go/variadic/multiple_return_add_sub
// origin: languages/go/tests/go/test_variadic.rs

package main
import "fmt"
func addSub(a int, b int) (int, int) { return a + b, a - b } var __buf string

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

func main() { s, d := addSub(10, 3)
__p(fmt.Sprint(s))
__p(fmt.Sprint(d))
__check("13\n7")
}
