// vybe-test: go/composite_literals_extra/array_literal_of_structs_runtime
// origin: languages/go/tests/go/test_composite_literals_extra.rs

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

func main() { pts := [2]point{{x: 1, y: 2}, {x: 3, y: 4}}
__p(fmt.Sprint(pts[1].y))
__check("4")
}
