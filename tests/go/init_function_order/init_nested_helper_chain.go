// vybe-test: go/init_function_order/init_nested_helper_chain
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var label string
func prefix() { label = "go" }
func suffix() { label += "lang" }
func init() { prefix() }
func init() { suffix() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(label), "golang") }
