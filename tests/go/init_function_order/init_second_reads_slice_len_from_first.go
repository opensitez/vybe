// vybe-test: go/init_function_order/init_second_reads_slice_len_from_first
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var items []string
var size int
func init() { items = []string{"a", "b"} }
func init() { size = len(items) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(size), "2") }
