// vybe-test: go/init_function_order/init_helper_increments_twice
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var count int
func bump() { count++ }
func init() { bump() }
func init() { bump() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(count), "2") }
