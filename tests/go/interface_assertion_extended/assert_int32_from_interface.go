// vybe-test: go/interface_assertion_extended/assert_int32_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = int32(-3)
n, ok := v.(int32)
__check(fmt.Sprint(n), "-3")
__check(fmt.Sprint(ok), "true") }
