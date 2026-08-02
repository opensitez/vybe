// vybe-test: go/init_function_order/init_closure_var_capture
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var fn func() int
var result int
func init() { n := 6
fn = func() int { return n } }
func init() { result = fn() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(result), "6") }
