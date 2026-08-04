// vybe-test: go/function_types_advanced/struct_passes_field_func_to_helper
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type box struct { transform func(int) int }
func invoke(b box, v int) int { return b.transform(v) }
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

func main() { value := box{transform: func(v int) int { return v - 1 }}
__p(fmt.Sprint(invoke(value, 9))) 
__check("8")
}
