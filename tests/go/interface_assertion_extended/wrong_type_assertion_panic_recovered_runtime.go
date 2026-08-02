// vybe-test: go/interface_assertion_extended/wrong_type_assertion_panic_recovered_runtime
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { __check(fmt.Sprint(recover() != nil), "true") }()
var v interface{} = 1
_ = v.(string) }
