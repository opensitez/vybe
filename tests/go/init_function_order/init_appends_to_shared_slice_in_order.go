// vybe-test: go/init_function_order/init_appends_to_shared_slice_in_order
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var seq []int
func init() { seq = append(seq, 10) }
func init() { seq = append(seq, 20) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(len(seq)), "2")
__check(fmt.Sprint(seq[0]), "10")
__check(fmt.Sprint(seq[1]), "20") }
