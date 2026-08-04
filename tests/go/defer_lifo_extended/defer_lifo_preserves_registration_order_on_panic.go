// vybe-test: go/defer_lifo_extended/defer_lifo_preserves_registration_order_on_panic
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer __p(fmt.Sprint(1))
defer __p(fmt.Sprint(2))
defer func() { recover() }()
panic("x") }
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

func main() { run() 
__check("2\n1")
}
