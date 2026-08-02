// vybe-test: go/interface_assertion_extended/single_assert_int_from_interface_runtime
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = 7
__check(fmt.Sprint(v.(int)), "7") }
