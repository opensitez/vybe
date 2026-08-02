// vybe-test: go/interface_nil_comparable/untyped_nil_assigned_to_empty_interface_is_nil
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value interface{} = nil
__check(fmt.Sprint(value == nil), "true") }
