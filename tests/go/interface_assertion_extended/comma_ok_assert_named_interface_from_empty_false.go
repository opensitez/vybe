// vybe-test: go/interface_assertion_extended/comma_ok_assert_named_interface_from_empty_false
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type speaker interface { talk() string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = 1
_, ok := v.(speaker)
__check(fmt.Sprint(ok), "false") }
