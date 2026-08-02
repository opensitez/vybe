// vybe-test: go/init_function_order/init_uses_const_defined_before
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
const base = 8
var scaled int
func init() { scaled = base }
func init() { scaled = scaled / 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(scaled), "4") }
