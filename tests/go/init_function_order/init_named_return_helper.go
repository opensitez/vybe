// vybe-test: go/init_function_order/init_named_return_helper
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var stored int
func read() (n int) { n = 9
return }
func init() { stored = read() }
func init() { stored++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(stored), "10") }
