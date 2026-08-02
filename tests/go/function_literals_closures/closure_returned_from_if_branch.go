// vybe-test: go/function_literals_closures/closure_returned_from_if_branch
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func pick(positive bool) func(int) int { if positive { return func(x int) int { return x } }
return func(x int) int { return -x } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(pick(false)(5)), "-5") }
