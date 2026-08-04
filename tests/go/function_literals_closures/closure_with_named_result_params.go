// vybe-test: go/function_literals_closures/closure_with_named_result_params
// origin: languages/go/tests/go/test_function_literals_closures.rs

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

func main() { divide := func(a, b int) (q int, r int) { q = a / b
r = a % b
return }
q, r := divide(10, 3)
__p(fmt.Sprint(q))
__p(fmt.Sprint(r)) 
__check("3\n1")
}
