// vybe-test: go/interface_assertion_extended/comma_ok_assert_string_from_empty_interface_true
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = "go"
s, ok := v.(string)
__check(fmt.Sprint(s), "go")
__check(fmt.Sprint(ok), "true") }
