// vybe-test: go/interface_assertion_extended/comma_ok_assert_error_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
import "errors"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = errors.New("e")
e, ok := v.(error)
__check(fmt.Sprint(e.Error()), "e")
__check(fmt.Sprint(ok), "true") }
