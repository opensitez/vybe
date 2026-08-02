// vybe-test: go/interface_assertion_extended/comma_ok_on_nil_interface_value
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{}
_, ok := v.(int)
__check(fmt.Sprint(ok), "false") }
