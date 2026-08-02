// vybe-test: go/interface_assertion_extended/untyped_nil_interface_is_nil
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
__check(fmt.Sprint(v == nil), "true") }
