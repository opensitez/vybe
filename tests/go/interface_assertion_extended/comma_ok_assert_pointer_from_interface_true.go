// vybe-test: go/interface_assertion_extended/comma_ok_assert_pointer_from_interface_true
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type node struct { id int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := &node{id: 5}
var v interface{} = n
p, ok := v.(*node)
__check(fmt.Sprint(p.id), "5")
__check(fmt.Sprint(ok), "true") }
