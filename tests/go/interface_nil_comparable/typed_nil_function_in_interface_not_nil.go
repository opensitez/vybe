// vybe-test: go/interface_nil_comparable/typed_nil_function_in_interface_not_nil
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var fn func()
var value interface{} = fn
__check(fmt.Sprint(value == nil), "false") }
