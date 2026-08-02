// vybe-test: go/interface_assertion_extended/assert_pointer_value_mismatch_comma_ok_false
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type node struct { v int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := node{v: 2}
var v interface{} = n
_, ok := v.(*node)
__check(fmt.Sprint(ok), "false") }
