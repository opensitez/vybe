// vybe-test: go/init_function_order/init_map_two_keys_sequential
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var m = map[int]string{}
func init() { m[1] = "one" }
func init() { m[2] = "two" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(len(m)), "2")
__check(fmt.Sprint(m[2]), "two") }
