// vybe-test: go/function_literals_closures/closure_capture_bool_toggle
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { on := false
flip := func() { on = !on }
flip()
flip()
__check(fmt.Sprint(on), "false") }
