// vybe-test: go/init_function_order/init_float64_product
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var ratio float64
func init() { ratio = 2.0 }
func init() { ratio = ratio * 1.5 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ratio), "3") }
