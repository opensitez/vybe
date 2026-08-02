// vybe-test: go/interface_assertion_extended/assert_interface_type_from_empty_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type fmtStringer interface { String() string }
type myInt int
func (m myInt) String() string { return "n" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = myInt(3)
s, ok := v.(fmtStringer)
__check(fmt.Sprint(s.String()), "n")
__check(fmt.Sprint(ok), "true") }
