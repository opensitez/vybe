// vybe-test: go/defer_panic_variants/defer_in_nested_loops_registers_six_callbacks
// origin: languages/go/tests/go/test_defer_panic_variants.rs

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

func main() { for i := 0; i < 2; i++ { for j := 0; j < 3; j++ { defer __p(fmt.Sprint(i*10 + j)) } } 
__check("12\n11\n10\n2\n1\n0")
}
