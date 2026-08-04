// vybe-test: go/function_literals_closures/closure_returned_from_if_branch
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func pick(positive bool) func(int) int { if positive { return func(x int) int { return x } }
return func(x int) int { return -x } }
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

func main() { __p(fmt.Sprint(pick(false)(5))) 
__check("-5")
}
