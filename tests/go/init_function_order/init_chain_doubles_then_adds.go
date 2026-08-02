// vybe-test: go/init_function_order/init_chain_doubles_then_adds
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var value int
func init() { value = 3 }
func init() { value = value * 2 }
func init() { value = value + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(value), "7") }
