// vybe-test: go/function_literals_closures/closure_as_interface_method
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
type runner interface { run() int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var r runner = runnerFunc(func() int { return 7 })
__check(fmt.Sprint(r.run()), "7") }
type runnerFunc func() int
func (f runnerFunc) run() int { return f() }
