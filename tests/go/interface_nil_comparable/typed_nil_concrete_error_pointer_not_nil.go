// vybe-test: go/interface_nil_comparable/typed_nil_concrete_error_pointer_not_nil
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
type myError struct { msg string }
func (e *myError) Error() string { return e.msg }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *myError
var err error = p
__check(fmt.Sprint(err == nil), "false") }
