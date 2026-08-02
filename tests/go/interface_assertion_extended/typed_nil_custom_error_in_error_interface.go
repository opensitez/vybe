// vybe-test: go/interface_assertion_extended/typed_nil_custom_error_in_error_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type e struct { msg string }
func (err *e) Error() string { return err.msg }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *e
var err error = p
__check(fmt.Sprint(err == nil), "false") }
