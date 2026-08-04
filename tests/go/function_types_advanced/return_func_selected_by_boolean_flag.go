// vybe-test: go/function_types_advanced/return_func_selected_by_boolean_flag
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func pick(positive bool) func(int) int { if positive { return func(v int) int { return v + 1 } }
return func(v int) int { return v - 1 } }
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

func main() { __p(fmt.Sprint(pick(true)(6)))
__p(fmt.Sprint(pick(false)(6))) 
__check("7\n5")
}
