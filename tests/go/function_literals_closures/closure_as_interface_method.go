// vybe-test: go/function_literals_closures/closure_as_interface_method
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
type runner interface { run() int }
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

func main() { var r runner = runnerFunc(func() int { return 7 })
__p(fmt.Sprint(r.run())) 
__check("7")
}
type runnerFunc func() int
func (f runnerFunc) run() int { return f() }
