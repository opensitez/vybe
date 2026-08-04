// vybe-test: go/pointer_receivers_advanced/new_assign_fields_equals_keyed_literal_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type point struct { x int
y int }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { fromNew := new(point)
fromNew.x = 3
fromNew.y = 8
fromLit := &point{x: 3, y: 8}
__p(fmt.Sprint(fromNew.x == fromLit.x))
__p(fmt.Sprint(fromNew.y == fromLit.y))
__check("true\ntrue")
}
