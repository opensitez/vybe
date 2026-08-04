// vybe-test: go/function_types_advanced/return_curried_multiplier_applied_twice
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func scale(factor int) func(int) int { return func(v int) int { return v * factor } }
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

func main() { double := scale(2)
triple := scale(3)
__p(fmt.Sprint(double(4)))
__p(fmt.Sprint(triple(4))) 
__check("8\n12")
}
