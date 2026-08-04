// vybe-test: go/functions_patterns_extra/function_literal_in_struct_field_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
type holder struct { fn func(int) int }
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

func main() { h := holder{fn: func(v int) int { return v * 3 }}
__p(fmt.Sprint(h.fn(4)))
__check("12")
}
