// vybe-test: go/init_function_order/init_named_return_helper
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var stored int
func read() (n int) { n = 9
return }
func init() { stored = read() }
func init() { stored++ }
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

func main() { __p(fmt.Sprint(stored)) 
__check("10")
}
