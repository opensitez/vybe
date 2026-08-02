// vybe-test: go/init_function_order/init_three_funcs_build_digit_string
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var digits string
func init() { digits = "1" }
func init() { digits += "2" }
func init() { digits += "3" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(digits), "123") }
