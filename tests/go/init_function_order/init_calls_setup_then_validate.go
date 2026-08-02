// vybe-test: go/init_function_order/init_calls_setup_then_validate
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var ok bool
func setup() { ok = true }
func validate() bool { return ok }
func init() { setup() }
func init() { ok = validate() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ok), "true") }
