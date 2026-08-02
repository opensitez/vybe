// vybe-test: go/function_literals_closures/closure_return_same_capture
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { base := 10
mk := func() func() int { return func() int { return base } }
__check(fmt.Sprint(mk()()), "10") }
