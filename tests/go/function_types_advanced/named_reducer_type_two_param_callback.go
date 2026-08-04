// vybe-test: go/function_types_advanced/named_reducer_type_two_param_callback
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type Reducer func(int, int) int
func fold(values []int, r Reducer, init int) int { acc := init
for _, v := range values { acc = r(acc, v) }
return acc }
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

func main() { __p(fmt.Sprint(fold([]int{1, 2, 3}, Reducer(func(a int, b int) int { return a + b }), 0))) 
__check("6")
}
