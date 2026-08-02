// vybe-test: go/interface_assertion_extended/two_typed_nil_different_types_not_equal
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var pi *int
var ps *string
var left interface{} = pi
var right interface{} = ps
__check(fmt.Sprint(left == right), "false") }
