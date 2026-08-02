// vybe-test: go/interface_assertion_extended/typed_nil_assert_to_concrete_pointer_ok
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type widget struct { n int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *widget
var v interface{} = p
_, ok := v.(*widget)
__check(fmt.Sprint(ok), "true") }
