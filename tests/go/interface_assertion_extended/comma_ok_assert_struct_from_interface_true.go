// vybe-test: go/interface_assertion_extended/comma_ok_assert_struct_from_interface_true
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type point struct { x int
y int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = point{x: 1, y: 2}
p, ok := v.(point)
__check(fmt.Sprint(p.x + p.y), "3")
__check(fmt.Sprint(ok), "true") }
