// vybe-test: go/function_types_advanced/func_returning_func_zero_value_nil
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var factory func(int) func(int) int
__check(fmt.Sprint(factory == nil), "true") }
