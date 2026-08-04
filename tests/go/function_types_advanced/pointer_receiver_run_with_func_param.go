// vybe-test: go/function_types_advanced/pointer_receiver_run_with_func_param
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type acc struct { total int }
func (a *acc) addEach(values []int, combine func(int, int) int) { for _, v := range values { a.total = combine(a.total, v) } }
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

func main() { value := acc{}
value.addEach([]int{1, 2, 3}, func(a int, b int) int { return a + b })
__p(fmt.Sprint(value.total)) 
__check("6")
}
