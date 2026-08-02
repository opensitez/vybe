// vybe-test: go/init_function_order/init_second_uses_first_bool_flag
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var armed bool
var ready bool
func init() { armed = true }
func init() { ready = armed }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ready), "true") }
