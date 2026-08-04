// vybe-test: go/function_literals_closures/return_closure_from_function
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func makeAdder(base int) func(int) int { return func(x int) int { return base + x } }
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

func main() { add5 := makeAdder(5)
__p(fmt.Sprint(add5(3))) 
__check("8")
}
