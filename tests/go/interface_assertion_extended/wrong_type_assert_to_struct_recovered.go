// vybe-test: go/interface_assertion_extended/wrong_type_assert_to_struct_recovered
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type a struct{}
type b struct{}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { __check(fmt.Sprint(recover() != nil), "true") }()
var v interface{} = a{}
_ = v.(b) }
