// vybe-test: go/higher_order_functions/function_returns_closure
// origin: languages/go/tests/go/test_higher_order_functions.rs

package main
import "fmt"
func adder(base int) func(int) int { return func(x int) int { return base + x } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { inc := adder(10)
__check(fmt.Sprint(inc(3)), "13") }
