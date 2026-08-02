// vybe-test: go/init_function_order/init_interface_value_assignment
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var holder interface{}
var tag string
func init() { holder = 42 }
func init() { tag = fmt.Sprint(holder) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(tag), "42") }
