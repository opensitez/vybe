// vybe-test: go/init_function_order/init_type_alias_conversion
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
type score int
var high score
func init() { high = score(11) }
func init() { high = high + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int(high)), "12") }
