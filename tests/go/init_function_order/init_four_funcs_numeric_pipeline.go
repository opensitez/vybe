// vybe-test: go/init_function_order/init_four_funcs_numeric_pipeline
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var n int
func init() { n = 2 }
func init() { n += 3 }
func init() { n *= 2 }
func init() { n -= 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(n), "9") }
