// vybe-test: go/init_function_order/init_three_call_same_helper
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var tally int
func add(n int) { tally += n }
func init() { add(1) }
func init() { add(2) }
func init() { add(3) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(tally), "6") }
