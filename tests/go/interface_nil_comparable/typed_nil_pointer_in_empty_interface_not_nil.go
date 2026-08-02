// vybe-test: go/interface_nil_comparable/typed_nil_pointer_in_empty_interface_not_nil
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *int
var value interface{} = p
__check(fmt.Sprint(value == nil), "false") }
