// vybe-test: go/interface_assertion_extended/comma_ok_assert_int_from_empty_interface_false
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = "x"
_, ok := v.(int)
__check(fmt.Sprint(ok), "false") }
