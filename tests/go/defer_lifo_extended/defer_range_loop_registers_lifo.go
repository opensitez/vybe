// vybe-test: go/defer_lifo_extended/defer_range_loop_registers_lifo
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
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

func main() { for _, v := range []int{10, 20} { defer __p(fmt.Sprint(v)) } 
__check("20\n10")
}
