// vybe-test: go/function_types_advanced/two_zero_value_funcs_same_type_equal
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left func()
var right func()
__check(fmt.Sprint(left == right), "true")
__check(fmt.Sprint(left == nil), "true") }
