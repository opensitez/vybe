// vybe-test: go/pointer_receivers_advanced/new_assign_fields_equals_keyed_literal_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

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

func main() { fromNew := new(point)
fromNew.x = 3
fromNew.y = 8
fromLit := &point{x: 3, y: 8}
__check(fmt.Sprint(fromNew.x == fromLit.x), "true")
__check(fmt.Sprint(fromNew.y == fromLit.y), "true")
}
