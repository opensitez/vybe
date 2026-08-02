// vybe-test: go/interface_assertion_extended/interface_nil_vs_typed_nil_equality
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var empty interface{}
var p *int
var typed interface{} = p
__check(fmt.Sprint(empty == typed), "false") }
