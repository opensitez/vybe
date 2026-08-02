// vybe-test: go/init_function_order/init_assigns_struct_then_reads_field
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
type pair struct { a int
b int }
var p pair
func init() { p = pair{a: 3, b: 4} }
func init() { p.b = p.a + p.b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(p.b), "7") }
