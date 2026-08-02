// vybe-test: go/init_function_order/init_uses_iota_const_group
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
const ( First = iota; Second; Third )
var picked int
func init() { picked = Second }
func init() { picked = picked + Third }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(picked), "3") }
