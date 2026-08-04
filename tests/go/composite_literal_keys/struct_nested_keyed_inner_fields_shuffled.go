// vybe-test: go/composite_literal_keys/struct_nested_keyed_inner_fields_shuffled
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type coord struct { x int
y int }
type rect struct { origin coord
size coord }
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

func main() { r := rect{size: coord{y: 4, x: 3}, origin: coord{x: 1, y: 2}}
__p(fmt.Sprint(r.origin.x))
__p(fmt.Sprint(r.size.y))
__check("1\n4")
}
