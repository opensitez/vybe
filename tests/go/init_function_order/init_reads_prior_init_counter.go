// vybe-test: go/init_function_order/init_reads_prior_init_counter
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var step int
func init() { step = 1 }
func init() { step = step * 10 }
func init() { step = step + 5 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(step), "15") }
