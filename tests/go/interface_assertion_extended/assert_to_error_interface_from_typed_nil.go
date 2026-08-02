// vybe-test: go/interface_assertion_extended/assert_to_error_interface_from_typed_nil
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type myErr struct { msg string }
func (e *myErr) Error() string { return e.msg }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *myErr
var err error = p
__check(fmt.Sprint(err == nil), "false")
_, ok := err.(*myErr)
__check(fmt.Sprint(ok), "true") }
