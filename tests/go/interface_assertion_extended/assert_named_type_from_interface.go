// vybe-test: go/interface_assertion_extended/assert_named_type_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type counter int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = counter(9)
c, ok := v.(counter)
__check(fmt.Sprint(int(c)), "9")
__check(fmt.Sprint(ok), "true") }
