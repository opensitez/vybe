// vybe-test: go/init_function_order/init_pointer_deref_chain
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var target int
var ptr *int
func init() { target = 5
ptr = &target }
func init() { *ptr = *ptr + 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(target), "7") }
