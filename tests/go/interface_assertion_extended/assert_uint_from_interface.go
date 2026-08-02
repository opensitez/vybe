// vybe-test: go/interface_assertion_extended/assert_uint_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = uint(12)
u, ok := v.(uint)
__check(fmt.Sprint(u), "12")
__check(fmt.Sprint(ok), "true") }
