// vybe-test: go/interface_assertion_extended/assert_float64_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = 2.5
f, ok := v.(float64)
__check(fmt.Sprint(f), "2.5")
__check(fmt.Sprint(ok), "true") }
