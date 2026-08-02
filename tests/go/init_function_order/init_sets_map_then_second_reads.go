// vybe-test: go/init_function_order/init_sets_map_then_second_reads
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var registry = map[string]int{}
var total int
func init() { registry["x"] = 4 }
func init() { total = registry["x"] + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(total), "5") }
