// vybe-test: go/init_function_order/init_bool_toggle_twice
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var flag bool
func init() { flag = !flag }
func init() { flag = !flag }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(flag), "false") }
