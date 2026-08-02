// vybe-test: go/interface_assertion_extended/assert_value_from_pointer_boxed_ok
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

func main() { n := &node{v: 2}
var v interface{} = n
p, ok := v.(*node)
__check(fmt.Sprint(p.v), "2")
__check(fmt.Sprint(ok), "true") }
