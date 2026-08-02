// vybe-test: go/interface_nil_comparable/interface_struct_inequality_different_fields
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

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

func main() { var left interface{} = point{x: 1, y: 2}
var right interface{} = point{x: 2, y: 1}
__check(fmt.Sprint(left == right), "false") }
