// vybe-test: go/defer_lifo_extended/defer_mixed_func_and_literal_order
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func tag(s string) { __p(fmt.Sprint(s)) }
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

func main() { defer tag("one")
defer func() { __p(fmt.Sprint("two")) }()
defer tag("three")
__check("three\ntwo\none")
}
