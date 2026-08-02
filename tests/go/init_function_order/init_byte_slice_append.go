// vybe-test: go/init_function_order/init_byte_slice_append
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var data []byte
func init() { data = append(data, 'x') }
func init() { data = append(data, 'y') }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(len(data)), "2")
__check(fmt.Sprint(int(data[1])), "121") }
