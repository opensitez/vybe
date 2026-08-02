// vybe-test: go/init_blank_import/init_order_three_sequential_appends
// origin: languages/go/tests/go/test_init_blank_import.rs

package main
import "fmt"
var order string
func init() { order = order + "1" }
func init() { order = order + "2" }
func init() { order = order + "3" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(order), "123") }
