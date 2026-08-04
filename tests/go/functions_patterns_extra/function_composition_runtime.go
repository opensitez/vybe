// vybe-test: go/functions_patterns_extra/function_composition_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func compose(a func(int) int, b func(int) int) func(int) int { return func(v int) int { return b(a(v)) } }
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

func main() { fn := compose(func(v int) int { return v + 1 }, func(v int) int { return v * 2 })
__p(fmt.Sprint(fn(5)))
__check("12")
}
