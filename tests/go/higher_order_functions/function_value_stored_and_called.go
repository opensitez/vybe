// vybe-test: go/higher_order_functions/function_value_stored_and_called
// origin: languages/go/tests/go/test_higher_order_functions.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var f func(int) int = func(x int) int { return x + 1 }
__check(fmt.Sprint(f(4)), "5") }
