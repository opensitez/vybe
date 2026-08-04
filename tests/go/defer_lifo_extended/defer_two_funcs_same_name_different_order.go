// vybe-test: go/defer_lifo_extended/defer_two_funcs_same_name_different_order
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func a() { __p(fmt.Sprint("a")) }
func b() { __p(fmt.Sprint("b")) }
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

func main() { defer a()
defer b()
__check("b\na")
}
