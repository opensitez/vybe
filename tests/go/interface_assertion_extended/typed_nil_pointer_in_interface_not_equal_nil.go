// vybe-test: go/interface_assertion_extended/typed_nil_pointer_in_interface_not_equal_nil
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *int
var v interface{} = p
__check(fmt.Sprint(v == nil), "false") }
