// vybe-test: go/defer_panic_variants/defer_lifo_mixed_named_funcs_and_literals
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func mark(label string) { __p(fmt.Sprint(label)) }
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

func main() { defer mark("alpha")
defer func() { __p(fmt.Sprint("beta")) }()
defer mark("gamma")
__check("gamma\nbeta\nalpha")
}
