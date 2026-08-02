// vybe-test: go/init_function_order/init_array_element_mutation
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var arr = [3]int{1, 1, 1}
func init() { arr[0] = 2 }
func init() { arr[1] = arr[0] + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(arr[0]), "2")
__check(fmt.Sprint(arr[1]), "3") }
