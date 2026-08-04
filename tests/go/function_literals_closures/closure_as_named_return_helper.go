// vybe-test: go/function_literals_closures/closure_as_named_return_helper
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func compute() int { transform := func(x int) int { return x + 10 }
return transform(5) }
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

func main() { __p(fmt.Sprint(compute())) 
__check("15")
}
