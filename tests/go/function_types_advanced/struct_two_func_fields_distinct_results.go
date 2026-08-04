// vybe-test: go/function_types_advanced/struct_two_func_fields_distinct_results
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type pair struct { left func() int
right func() int }
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

func main() { value := pair{left: func() int { return 1 }, right: func() int { return 2 }}
__p(fmt.Sprint(value.left()))
__p(fmt.Sprint(value.right())) 
__check("1\n2")
}
