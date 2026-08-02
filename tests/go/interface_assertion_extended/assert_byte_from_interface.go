// vybe-test: go/interface_assertion_extended/assert_byte_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = byte(65)
b, ok := v.(byte)
__check(fmt.Sprint(b), "65")
__check(fmt.Sprint(ok), "true") }
