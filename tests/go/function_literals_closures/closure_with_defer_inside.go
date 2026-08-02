// vybe-test: go/function_literals_closures/closure_with_defer_inside
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run := func() { defer __check(fmt.Sprint("done"), "go")
__check(fmt.Sprint("go"), "done") }
run() }
