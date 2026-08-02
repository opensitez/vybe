// vybe-test: go/interface_assertion_extended/two_typed_nil_same_type_equal_in_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var a *int
var b *int
var left interface{} = a
var right interface{} = b
__check(fmt.Sprint(left == right), "true") }
