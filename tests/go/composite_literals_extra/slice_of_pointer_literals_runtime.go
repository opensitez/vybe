// vybe-test: go/composite_literals_extra/slice_of_pointer_literals_runtime
// origin: languages/go/tests/go/test_composite_literals_extra.rs

package main
import "fmt"
type point struct { x int }
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

func main() { a := &point{x: 1}
b := &point{x: 2}
values := []*point{a, b}
__p(fmt.Sprint(values[1].x))
__check("2")
}
