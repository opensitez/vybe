// vybe-test: go/interface_nil_comparable/interface_reset_to_nil_after_value
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value interface{} = 9
__check(fmt.Sprint(value == nil), "false")
value = nil
__check(fmt.Sprint(value == nil), "true") }
