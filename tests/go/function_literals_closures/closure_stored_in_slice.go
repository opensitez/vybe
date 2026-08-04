// vybe-test: go/function_literals_closures/closure_stored_in_slice
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

func main() { fns := []func(int) int{ func(x int) int { return x + 1 }, func(x int) int { return x + 2 } }
__p(fmt.Sprint(fns[0](5)))
__p(fmt.Sprint(fns[1](5))) 
__check("6\n7")
}
