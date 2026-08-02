// vybe-test: go/init_function_order/init_two_funcs_append_letters
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var trace string
func init() { trace += "A" }
func init() { trace += "B" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(trace), "AB") }
